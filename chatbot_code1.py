#!/usr/bin/env python
# coding: utf-8

import requests
import pandas as pd
import re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import calendar
import textwrap
import os
import io

# ── Word doc export (optional — only if python-docx is installed) ──
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

engine = create_engine('postgresql://postgres:11223344@localhost:5432/sales_db')

print()
print("=" * 60)
print("       SUPERSTORE SALES DATABASE CHATBOT")
print("=" * 60)
print()

try:
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print("  ✅ Database connected successfully")
except Exception as e:
    print(f"  ❌ Database connection failed: {e}")

try:
    requests.post(
        url="http://192.168.15.220:11434/api/generate",
        headers={"Content-Type": "application/json"},
        json={
            "model": "gemma3:12b",
            "prompt": "hi",
            "stream": False,
            "options": {
                "temperature": 0.1,
                "num_ctx": 8192
            }
        },
        timeout=60
    )
    print("  ✅ AI model connected successfully (gemma3:12b)")
except Exception as e:
    print(f"  ❌ AI model connection failed: {e}")

if DOCX_AVAILABLE:
    print("  ✅ Word export available (python-docx installed)")
else:
    print("  ⚠️  Word export unavailable — run: pip install python-docx")

print()
print("=" * 60)

context = []
latest_chart_buf = None

def get_metadata():
    metadata = "DATABASE STRUCTURE:\n"
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name 
                FROM information_schema.tables 
                WHERE table_schema = 'public'
                ORDER BY table_name
            """)).fetchall()
            for table in tables:
                table_name = table[0]
                metadata += f"\nTable: {table_name}\n"
                metadata += "Columns:\n"
                columns = conn.execute(text(f"""
                    SELECT column_name, data_type 
                    FROM information_schema.columns 
                    WHERE table_name = '{table_name}'
                    ORDER BY ordinal_position
                """)).fetchall()
                for col in columns:
                    metadata += f"  - {col[0]} ({col[1]})\n"
                count = conn.execute(text(f"SELECT COUNT(*) FROM {table_name}")).fetchone()[0]
                metadata += f"Total rows: {count}\n"
    except Exception as e:
        print(f"  ❌ Metadata error: {e}")
    return metadata

print("\n  📋 Database metadata loaded:")
metadata = get_metadata()
print(metadata)
print("=" * 60)

def get_special_columns():
    special_cols = {}
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name, column_name
                FROM information_schema.columns
                WHERE table_schema = 'public'
                AND column_name ~ '[^a-zA-Z0-9_]'
            """)).fetchall()
            for table_name, col_name in tables:
                normalized = col_name.replace('-', '_').replace(' ', '_').replace('.', '_')
                special_cols[normalized] = f'"{col_name}"'
                special_cols[col_name] = f'"{col_name}"'
    except Exception as e:
        print(f"  ⚠️ Could not get special columns: {e}")
    return special_cols

special_columns = get_special_columns()
print(f"  ✅ Special columns detected: {special_columns}")

def get_table_columns():
    table_cols = {}
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name, column_name
                FROM information_schema.columns
                WHERE table_schema = 'public'
                ORDER BY table_name, ordinal_position
            """)).fetchall()
            for table_name, col_name in tables:
                if table_name not in table_cols:
                    table_cols[table_name] = []
                table_cols[table_name].append(col_name.lower())
    except Exception as e:
        print(f"  ⚠️ Could not get table columns: {e}")
    return table_cols

table_columns = get_table_columns()
print(f"  ✅ Table columns loaded: {list(table_columns.keys())}")


# ══════════════════════════════════════════════════════════════
# DATE INJECTION SYSTEM
# ══════════════════════════════════════════════════════════════

MONTH_NAMES = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8,
    'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,
    'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9,
    'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

QUARTER_MONTHS = {1: 1, 2: 4, 3: 7, 4: 10}

def get_quarter_range(year, quarter):
    start_month = QUARTER_MONTHS[quarter]
    start = datetime(year, start_month, 1)
    if quarter == 4:
        end = datetime(year + 1, 1, 1)
    else:
        end = datetime(year, QUARTER_MONTHS[quarter + 1], 1)
    return start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')

def detect_date_context(question):
    q = question.lower().strip()
    today = datetime.now()
    now_year = today.year
    now_month = today.month
    result = None

    m = re.search(r'\b(\d{1,2})(?:st|nd|rd|th)?\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
    if m:
        day, month_str, year = int(m.group(1)), m.group(2), int(m.group(3))
        month = MONTH_NAMES[month_str]
        date_str = f"{year}-{month:02d}-{day:02d}"
        result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{4})\b', q)
        if m:
            month_str, day, year = m.group(1), int(m.group(2)), int(m.group(3))
            month = MONTH_NAMES[month_str]
            date_str = f"{year}-{month:02d}-{day:02d}"
            result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(\d{4}-\d{2}-\d{2})\b', q)
        if m:
            date_str = m.group(1)
            result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
        if m:
            month_str, year = m.group(1), int(m.group(2))
            month = MONTH_NAMES[month_str]
            start = f"{year}-{month:02d}-01"
            end_month = month + 1 if month < 12 else 1
            end_year = year if month < 12 else year + 1
            end = f"{end_year}-{end_month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"in {m.group(1).title()} {year}"}

    if not result:
        m = re.search(r'\bq([1-4])\s+(?:of\s+)?(\d{4})\b', q)
        if m:
            quarter, year = int(m.group(1)), int(m.group(2))
            start, end = get_quarter_range(year, quarter)
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"Q{quarter} {year}"}

    if not result:
        quarter_words = {'first': 1, 'second': 2, 'third': 3, 'fourth': 4}
        m = re.search(r'\b(first|second|third|fourth)\s+quarter\s+(?:of\s+)?(\d{4})\b', q)
        if m:
            quarter, year = quarter_words[m.group(1)], int(m.group(2))
            start, end = get_quarter_range(year, quarter)
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"Q{quarter} {year}"}

    if not result:
        m = re.search(r'\b(in|during|for|of)\s+(202[0-9])\b', q)
        if not m:
            m = re.search(r'\b(202[0-9])\b', q)
            if m:
                year = int(m.group(1))
                result = {'filter': f"order_date >= '{year}-01-01'::date AND order_date < '{year+1}-01-01'::date", 'desc': f"in {year}"}
        else:
            year = int(m.group(2))
            result = {'filter': f"order_date >= '{year}-01-01'::date AND order_date < '{year+1}-01-01'::date", 'desc': f"in {year}"}

    if not result and re.search(r'\blast\s+week\b', q):
        start = (today - timedelta(days=7)).strftime('%Y-%m-%d')
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "last week"}

    if not result and re.search(r'\bthis\s+week\b', q):
        start_of_week = (today - timedelta(days=today.weekday())).strftime('%Y-%m-%d')
        result = {'filter': f"order_date >= '{start_of_week}'::date AND order_date < CURRENT_DATE", 'desc': "this week"}

    if not result and re.search(r'\blast\s+month\b', q):
        if now_month == 1:
            start = f"{now_year - 1}-12-01"
            end = f"{now_year}-01-01"
        else:
            start = f"{now_year}-{now_month-1:02d}-01"
            end = f"{now_year}-{now_month:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last month"}

    if not result and re.search(r'\bmonth\s+before\s+(last|that)\b|\b2\s+months\s+ago\b', q):
        two_months_ago = now_month - 2
        year_adj = now_year
        if two_months_ago <= 0:
            two_months_ago += 12
            year_adj -= 1
        one_month_ago = now_month - 1
        year_adj2 = now_year
        if one_month_ago <= 0:
            one_month_ago += 12
            year_adj2 -= 1
        start = f"{year_adj}-{two_months_ago:02d}-01"
        end = f"{year_adj2}-{one_month_ago:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "month before last"}

    if not result and re.search(r'\bthis\s+month\b', q):
        start = f"{now_year}-{now_month:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this month"}

    if not result and re.search(r'\blast\s+quarter\b', q):
        current_quarter = (now_month - 1) // 3 + 1
        last_quarter = current_quarter - 1
        year = now_year
        if last_quarter == 0:
            last_quarter = 4
            year -= 1
        start, end = get_quarter_range(year, last_quarter)
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last quarter"}

    if not result and re.search(r'\bthis\s+quarter\b', q):
        current_quarter = (now_month - 1) // 3 + 1
        start, end = get_quarter_range(now_year, current_quarter)
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this quarter"}

    if not result and re.search(r'\blast\s+year\b', q):
        start = f"{now_year - 1}-01-01"
        end = f"{now_year}-01-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last year"}

    if not result and re.search(r'\bthis\s+year\b', q):
        start = f"{now_year}-01-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this year"}

    if not result:
        m = re.search(r'\blast\s+(\d+)\s+days?\b', q)
        if m:
            n = int(m.group(1))
            start = (today - timedelta(days=n)).strftime('%Y-%m-%d')
            result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': f"last {n} days"}

    if not result:
        m = re.search(r'\blast\s+(\d+)\s+months?\b', q)
        if m:
            n = int(m.group(1))
            month = now_month - n
            year = now_year
            while month <= 0:
                month += 12
                year -= 1
            start = f"{year}-{month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': f"last {n} months"}

    if not result:
        m = re.search(r'\bbetween\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\s+and\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
        if m:
            m1, y1 = MONTH_NAMES[m.group(1)], int(m.group(2))
            m2, y2 = MONTH_NAMES[m.group(3)], int(m.group(4))
            start = f"{y1}-{m1:02d}-01"
            end_month = m2 + 1 if m2 < 12 else 1
            end_year = y2 if m2 < 12 else y2 + 1
            end = f"{end_year}-{end_month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"between {m.group(1).title()} {y1} and {m.group(3).title()} {y2}"}

    return result


def inject_date_into_prompt(question):
    comparison_words = [
        "compare", "vs", "versus", "month before", "last month vs",
        "this year vs", "last year vs", "previous", "month before last",
        "2 months ago", "last quarter vs", "this quarter vs"
    ]
    if any(word in question.lower() for word in comparison_words):
        print(f"  📅 Comparison question — skipping date injection")
        return None

    date_ctx = detect_date_context(question)
    if date_ctx:
        print(f"  📅 Date detected: {date_ctx['desc']} → {date_ctx['filter']}")
        return date_ctx
    return None


def chat_with_user(question):
    prompt = f"""You are a friendly and helpful Sales Analytics Assistant for a retail company called Superstore.
You are having a casual conversation with a business user.
Today's date is {datetime.now().strftime('%Y-%m-%d')}.

The user said: '{question}'

Rules:
- Respond naturally and conversationally like a helpful colleague
- Keep response strictly under 2 sentences and dont extend it too much.
- If they greet you, greet them back warmly and offer to help
- If they say thanks, acknowledge it nicely and offer more help
- If they ask what you can do, briefly mention sales analytics capabilities
- If they seem confused, offer 2-3 example questions they can ask
- If they ask something completely outside sales data, politely redirect
- Never mention SQL or technical terms
- Sound human and friendly, not robotic
- Never start with 'I'"""

    response = requests.post(
        url="http://192.168.15.220:11434/api/generate",
        headers={"Content-Type": "application/json"},
        json={
            "model": "gemma3:12b",
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.8,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
                "num_ctx": 2048
            }
        },
        timeout=60
    )
    return response.json()['response'].strip()


def save_to_word(context_history):
    if not DOCX_AVAILABLE:
        return "  ⚠️  python-docx not installed. Run: pip install python-docx --break-system-packages"

    if not context_history:
        return "  ⚠️  No conversation to save yet. Ask some questions first!"

    try:
        doc = Document()

        title = doc.add_heading('Superstore Sales Chat Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_para = doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        total_para = doc.add_paragraph(f'Total Questions: {len(context_history)}')
        total_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph('')

        doc.add_heading('Chat Summary', level=1)
        summary_table = doc.add_table(rows=1, cols=2)
        summary_table.style = 'Light Grid Accent 1'
        hdr = summary_table.rows[0].cells
        hdr[0].text = 'Question'
        hdr[1].text = 'Answer'
        for cell in hdr:
            cell.paragraphs[0].runs[0].bold = True

        for exchange in context_history:
            row = summary_table.add_row().cells
            row[0].text = exchange['question']
            clean_answer = exchange['answer'].split('\n\n  💡')[0].strip()
            row[1].text = clean_answer

        doc.add_paragraph('')
        doc.add_heading('Detailed Analysis', level=1)

        for i, exchange in enumerate(context_history):
            doc.add_heading(f'Q{i+1}: {exchange["question"]}', level=2)
            clean_answer = exchange['answer'].split('\n\n  💡')[0].strip()
            doc.add_paragraph(clean_answer)

            sql_para = doc.add_paragraph()
            sql_run = sql_para.add_run(f'SQL: {exchange["sql"]}')
            sql_run.font.size = Pt(8)
            sql_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

            try:
                result = run_query(exchange['sql'], exchange['question'])
                if result is not None and not result.empty and len(result) <= 50:
                    doc.add_paragraph('Data:').runs[0].bold = True
                    tbl = doc.add_table(rows=1, cols=len(result.columns))
                    tbl.style = 'Light Grid Accent 1'
                    hdr_cells = tbl.rows[0].cells
                    for j, col in enumerate(result.columns):
                        hdr_cells[j].text = str(col).replace('_', ' ').title()
                        hdr_cells[j].paragraphs[0].runs[0].bold = True
                    for _, row_data in result.head(20).iterrows():
                        cells = tbl.add_row().cells
                        for j, val in enumerate(row_data):
                            cells[j].text = str(round(val, 2) if isinstance(val, float) else val)
            except Exception:
                pass

            doc.add_paragraph('─' * 60)

        filename = f"chat_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        doc.save(filename)
        return f"  📄 Full chat report saved as: {filename}  ({len(context_history)} questions included)"

    except Exception as e:
        return f"  ❌ Could not save Word doc: {e}"


def needs_visualisation(question):
    no_chart_phrases = [
        "give me details", "show me details", "tell me about",
        "what is this database", "how can you help", "what are you",
        "what do you do", "how many unique", "what is the total",
        "who is the", "what is the name", "what is this",
        "which customer", "which product", "which state",
        "which city", "which region", "which category"
    ]
    question_lower = question.lower()
    if any(phrase in question_lower for phrase in no_chart_phrases):
        return False

    always_chart_phrases = [
        "by region", "by category", "by segment",
        "by state", "by city", "by month", "by quarter", "by year",
        "trend", "compare", "comparison", "breakdown", "distribution",
        "percentage", "percent", "chart", "graph", "plot",
        "visualise", "visualize", "monthly", "quarterly",
        "yearly", "growth", "pattern", "refund",
        "sales by", "profit by", "orders by",
        "return rate by", "return reasons", "top ", "bottom ",
        "show chart", "with chart", "vs", "versus",
        "cumulative", "running total"
    ]
    if any(phrase in question_lower for phrase in always_chart_phrases):
        return True

    return False


def create_visualisation(result, question):
    global latest_chart_buf  # declared once at top of function
    try:
        if len(result.columns) < 2:
            return None

        if len(result) == 1 and len(result.columns) == 2:
            col1, col2 = result.columns[0], result.columns[1]
            val1 = pd.to_numeric(result[col1].iloc[0], errors='coerce')
            val2 = pd.to_numeric(result[col2].iloc[0], errors='coerce')
            if pd.notna(val1) and pd.notna(val2):
                reshaped = pd.DataFrame({
                    'period': [col1.replace('_', ' ').title(), col2.replace('_', ' ').title()],
                    'value': [val1, val2]
                })
                result = reshaped

        if len(result) < 2:
            return None

        fig, ax = plt.subplots(figsize=(12, 6))
        fig.patch.set_facecolor('#1c1915')
        ax.set_facecolor('#1c1915')
        x_col = result.columns[0]
        y_col = result.columns[1]
        x_values = result[x_col].astype(str)
        y_values = pd.to_numeric(result[y_col], errors='coerce')
        question_lower = question.lower()

        def format_x_label(val):
            try:
                if re.match(r'^\d{4}-\d{2}$', val):
                    dt = datetime.strptime(val, '%Y-%m')
                    return dt.strftime('%b %Y')
                return val
            except Exception:
                return val

        x_values = x_values.apply(format_x_label)

        # PIE CHART
        if any(word in question_lower for word in ["pie", "percentage", "percent", "distribution", "contribution"]):
            numeric_col = None
            for col in result.columns:
                if pd.to_numeric(result[col], errors='coerce').notna().all():
                    numeric_col = col
                    break
            if numeric_col is None:
                return None
            if len(result) == 1:
                pct = min(float(result[numeric_col].iloc[0]), 100)
                values = [pct, 100 - pct]
                labels = [f'Yes ({pct:.1f}%)', f'No ({100-pct:.1f}%)']
                colors = ['#E85D04', '#1D9E75']
            else:
                values = pd.to_numeric(result[numeric_col], errors='coerce')
                labels = result[result.columns[0]].astype(str)
                colors = ['#2E75B6', '#1D9E75', '#E85D04', '#534AB7', '#BA7517',
                         '#C62828', '#00838F', '#558B2F', '#6A1B9A', '#2E7D32']
            wedges, texts, autotexts = ax.pie(
                values, labels=labels, colors=colors[:len(values)],
                autopct='%1.1f%%', startangle=90,
                wedgeprops={'edgecolor': 'white', 'linewidth': 2}
            )
            ax.set_title(question.title(), fontsize=14, fontweight='bold', pad=20, color='#e2e8f0')
            plt.tight_layout()
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buf.seek(0)
            latest_chart_buf = buf
            print(f"\n  📊 Pie chart generated!")
            return True

        # AREA CHART — before line chart
        if any(word in question_lower for word in ["cumulative", "running total", "accumulated"]):
            ax.fill_between(range(len(x_values)), y_values, alpha=0.4, color='#2E75B6')
            ax.plot(range(len(x_values)), y_values, color='#2E75B6', linewidth=2)
            ax.set_xticks(range(len(x_values)))
            ax.set_xticklabels(x_values, rotation=45, ha='right')
            chart_type = "Area"

        # LINE CHART
        elif any(word in question_lower for word in ["trend", "monthly", "quarterly", "yearly", "over time", "by month", "by year", "by quarter"]):
            ax.plot(x_values, y_values, marker='o', linewidth=2.5,
                   color='#2E75B6', markersize=8, markerfacecolor='white', markeredgewidth=2)
            ax.fill_between(range(len(x_values)), y_values, alpha=0.1, color='#2E75B6')
            chart_type = "Line"

        # SCATTER CHART
        elif any(word in question_lower for word in ["correlation", "relationship", "scatter"]):
            ax.scatter(x_values, y_values, color='#2E75B6', alpha=0.6, s=100, edgecolors='white')
            chart_type = "Scatter"

        # GROUPED BAR
        elif any(word in question_lower for word in ["vs", "versus", "compare", "comparison"]) and len(result.columns) >= 3:
            x = range(len(result))
            width = 0.35
            col_colors = ['#2E75B6', '#E85D04', '#1D9E75', '#534AB7']
            for i, col in enumerate(result.columns[1:3]):
                vals = pd.to_numeric(result[col], errors='coerce')
                offset = (i - 0.5) * width
                bars = ax.bar([xi + offset for xi in x], vals, width,
                             label=col.replace('_', ' ').title(),
                             color=col_colors[i], edgecolor='white')
                for bar, val in zip(bars, vals):
                    ax.text(bar.get_x() + bar.get_width()/2., bar.get_height(),
                           f'{val:,.0f}', ha='center', va='bottom', fontsize=8)
            ax.set_xticks(list(x))
            ax.set_xticklabels(result.iloc[:, 0].astype(str), rotation=45, ha='right')
            ax.legend()
            chart_type = "Grouped Bar"

        # HORIZONTAL BAR
        elif any(word in question_lower for word in ["top", "bottom"]) and len(result) > 3:
            sorted_df = result.copy()
            sorted_df[y_col] = pd.to_numeric(sorted_df[y_col], errors='coerce')
            sorted_df = sorted_df.sort_values(by=y_col, ascending=True)
            x_values = sorted_df[x_col].astype(str).apply(format_x_label)
            y_values = sorted_df[y_col]
            colors = plt.cm.Blues([0.3 + 0.7 * i / len(sorted_df) for i in range(len(sorted_df))])
            bars = ax.barh(x_values, y_values, color=colors, edgecolor='white')
            for bar, val in zip(bars, y_values):
                ax.text(bar.get_width(), bar.get_y() + bar.get_height()/2.,
                       f' {val:,.2f}', ha='left', va='center', fontsize=9, fontweight='bold',color='#e2e8f0')
            chart_type = "Horizontal Bar"

        # VERTICAL BAR (default)
        else:
            colors = ['#2E75B6', '#1D9E75', '#E85D04', '#534AB7', '#BA7517',
                     '#C62828', '#00838F', '#558B2F', '#6A1B9A', '#2E7D32']
            bars = ax.bar(x_values, y_values, color=colors[:len(result)],
                         edgecolor='white', linewidth=0.5, width=0.6)
            for bar, val in zip(bars, y_values):
                ax.text(bar.get_x() + bar.get_width()/2., bar.get_height(),
                       f'{val:,.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#e2e8f0')
            chart_type = "Bar"

        
        ax.set_title(question.title(), fontsize=14, fontweight='bold', pad=20, color='#e2e8f0')
        if chart_type == "Horizontal Bar":
            ax.set_xlabel(y_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.set_ylabel(x_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
        else:
            ax.set_xlabel(x_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.set_ylabel(y_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.grid(axis='y', alpha=0.15, linestyle='--', color='white')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#333333')
        ax.spines['bottom'].set_color('#333333')        
        ax.tick_params(colors='#aaaaaa')
        plt.xticks(rotation=45, ha='right', color='#aaaaaa')
        plt.yticks(color='#aaaaaa')
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        latest_chart_buf = buf
        print(f"\n  📊 {chart_type} chart generated!")
        return True

    except Exception as e:
        print(f"  ⚠️ Could not generate chart: {e}")
        return None


def text_to_sql(question, date_ctx=None):
    context_str = ""
    if context:
        context_str = "\n\nCONVERSATION HISTORY:\n"
        for exchange in context:
            context_str += f"User: {exchange['question']}\n"
            context_str += f"SQL: {exchange['sql']}\n"
            context_str += f"Answer: {exchange['answer']}\n"
            context_str += f"Result data: {exchange.get('result', '')}\n\n"
        context_str += """CRITICAL FOLLOW UP RULES:
- Read the conversation history AND result data above very carefully
- The result data shows the EXACT values returned by the database
- When user says 'that region' or 'it' or 'its' — look at the previous answer to find the exact region name and add WHERE region = 'exact name'
- When user says 'that customer' — add WHERE customer_name = 'exact name from previous answer'
- When user says 'that category' — add WHERE category = 'exact name from previous answer'
- When user says 'that city' — add WHERE city = 'exact name from previous answer'
- When user says 'that state' — add WHERE state = 'exact name from previous answer'
- When user says 'what is the profit' after asking about a specific region — use WHERE region = 'that specific region' not ORDER BY profit
- When user says 'how many orders does it have' after asking about a region — use WHERE region = 'that region' not WHERE customer_name
- NEVER use WHERE customer_name when the previous question was about a region
- NEVER use WHERE region when the previous question was about a customer
- Always match the filter column to what was being discussed in the previous question
- For 'which has the lowest' after a highest question — use same GROUP BY and metric but ORDER BY ASC LIMIT 1
- Never return UNRELATED when there is conversation history\n"""

    date_instruction = ""
    if date_ctx:
        date_instruction = f"""
CRITICAL DATE INSTRUCTION:
The user is asking about: {date_ctx['desc']}
Use EXACTLY this date filter in your WHERE clause (already calculated for you):
  {date_ctx['filter']}
Do NOT calculate dates yourself — just use the filter above as-is.
"""

    prompt = f"""You are an expert Sales Analytics Assistant for a retail company. You have direct access to the company's PostgreSQL sales database and your job is to help business users get insights from the data by answering their questions clearly and professionally.

You are speaking to business users who do not know SQL. They will ask questions in plain English and you must convert them into valid PostgreSQL SQL queries.

Today's date is {datetime.now().strftime('%Y-%m-%d')}.

{metadata}

The orders table has data from 2022 to 2026.
The returns table is linked to orders via order_id.

MOST IMPORTANT RULE:
- If the question is unrelated to the database return exactly: UNRELATED
- NEVER return SQL comments (lines starting with --)
- NEVER return explanations instead of SQL
- NEVER say you need more information — always use CURRENT_DATE for date calculations
- If you are unsure about a date range, use CURRENT_DATE and calculate it yourself

{date_instruction}

RELATIONSHIP BETWEEN TABLES:
- orders and returns are linked via order_id
- Always read the DATABASE STRUCTURE above carefully before writing SQL
- Only use columns that actually exist in each table as listed in DATABASE STRUCTURE
- For return rate by any group always use this exact pattern:
  SELECT o.group_column, ROUND((COUNT(DISTINCT r.return_id)::numeric / NULLIF(COUNT(DISTINCT o.order_id)::numeric, 0)) * 100, 2) AS return_rate FROM orders o LEFT JOIN returns r ON r.order_id = o.order_id GROUP BY o.group_column ORDER BY return_rate DESC
- Always prefix ambiguous columns with table alias: o.customer_name not just customer_name
- Never write correlated subqueries for return rates — always use LEFT JOIN with COUNT
- Never reference a column in a table that does not have it according to DATABASE STRUCTURE
- Use JOIN when question needs data from both tables
- The returns table does NOT have region, state, city, segment, category, customer_name or product columns
- To get returns by region: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.region
- To get returns by category: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.category
- To get returns by segment: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.segment
- To get returns by state: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.state
- To get returns by city: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.city
- To get returns by customer: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.customer_name
- Any time user asks for returns breakdown BY any group — always JOIN returns with orders

BASIC SQL RULES:
- Return ONLY the raw SQL query, no backticks, no markdown, no explanations, no semicolon
- Always include FROM clause in every query
- Always include every non-aggregate column in GROUP BY
- Always use ROUND(SUM(column)::numeric, 2) for decimal results
- Never nest aggregate functions like SUM(SUM()) or ROUND(ROUND())
- Always alias columns with meaningful names using AS
- Always include the GROUP BY column in SELECT
- For states always use WHERE state = 'StateName' never WHERE country = 'StateName'
- highest maximum top most all mean ORDER BY DESC
- lowest minimum bottom least all mean ORDER BY ASC
- For product searches use ILIKE with wildcards: WHERE product_name ILIKE '%name%'
- When user asks for sales always use SUM(sales) never SUM(profit)
- When user asks for profit always use SUM(profit) never SUM(sales)
- Never use WHERE customer_name ILIKE '%name%' unless user gives a real specific name
- Always use lowercase table names
- Never use COUNT(*) always use COUNT(DISTINCT column_name)
- When using JOIN always prefix ALL column references with table alias to avoid ambiguity
- For average discount: SELECT ROUND(AVG(discount)::numeric, 4) AS avg_discount FROM orders
- Never multiply discount by 100 in SQL, the post-processing will handle it

COUNTING RULES:
- For counting orders: COUNT(DISTINCT order_id)
- For counting customers: COUNT(DISTINCT customer_name)
- Always use DISTINCT when counting to avoid duplicates
- For which customer has most orders: SELECT customer_name, COUNT(DISTINCT order_id) AS total_orders FROM orders GROUP BY customer_name ORDER BY total_orders DESC LIMIT 1

AGGREGATION RULES:
- For top N: ORDER BY value DESC LIMIT N
- For bottom N: ORDER BY value ASC LIMIT N
- For profit margin: ROUND((SUM(profit)::numeric / NULLIF(SUM(sales)::numeric, 0)) * 100, 2)
- For average order value: ROUND((SUM(sales)::numeric / NULLIF(COUNT(DISTINCT order_id), 0)), 2)
- For percentage of total use subquery: ROUND(((SUM(sales)::numeric / NULLIF((SELECT SUM(sales)::numeric FROM orders), 0)) * 100)::numeric, 2)
- For overall return rate use: SELECT ROUND(((SELECT COUNT(DISTINCT return_id) FROM returns)::numeric / NULLIF((SELECT COUNT(DISTINCT order_id) FROM orders)::numeric, 0)) * 100, 2) AS return_percentage
- For loss making orders: WHERE profit < 0
- For high discount orders: WHERE discount > 0.30
- For cumulative/running total by month use window function:
  SELECT TO_CHAR(order_date, 'YYYY-MM') AS year_month,
  ROUND(SUM(SUM(sales)) OVER (ORDER BY TO_CHAR(order_date, 'YYYY-MM'))::numeric, 2) AS cumulative_sales
  FROM orders
  GROUP BY TO_CHAR(order_date, 'YYYY-MM')
  ORDER BY year_month
- NEVER use regular SUM for cumulative — always use SUM(SUM()) OVER() window function
- SUM(SUM(column)) OVER() is the correct pattern — outer SUM is the window, inner SUM is the aggregate

DELIVERY RULES:
- For delivery days: EXTRACT(EPOCH FROM (ship_date - order_date)) / 86400
- For average delivery time: SELECT ROUND(AVG(EXTRACT(EPOCH FROM (ship_date - order_date)) / 86400)::numeric, 2) AS avg_delivery_days FROM orders
- For late deliveries: WHERE EXTRACT(EPOCH FROM (ship_date - order_date)) / 86400 > 5

DATE RULES:
- Today is {datetime.now().strftime('%Y-%m-%d')}. You ALWAYS know the current date. NEVER say you don't know the date.
- NEVER return a comment or explanation. ALWAYS return a valid SQL query.
- Always use DATE_TRUNC for date range comparisons, never EXTRACT(MONTH) or EXTRACT(YEAR) alone
- For last month: WHERE order_date >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 months') AND order_date < DATE_TRUNC('month', CURRENT_DATE)
- For month before last: WHERE order_date >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months') AND order_date < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 months')
- For last quarter: WHERE order_date >= DATE_TRUNC('quarter', CURRENT_DATE - INTERVAL '3 months') AND order_date < DATE_TRUNC('quarter', CURRENT_DATE)
- For last year: WHERE order_date >= DATE_TRUNC('year', CURRENT_DATE - INTERVAL '1 years') AND order_date < DATE_TRUNC('year', CURRENT_DATE)
- For last week: WHERE order_date >= CURRENT_DATE - INTERVAL '7 days'
- For this year: WHERE order_date >= DATE_TRUNC('year', CURRENT_DATE)
- For this month: WHERE order_date >= DATE_TRUNC('month', CURRENT_DATE)
- For a specific year: WHERE order_date >= '2024-01-01'::date AND order_date < '2025-01-01'::date
- For a specific date: WHERE order_date = '2024-03-30'::date
- For comparing two periods always use CASE WHEN with DATE_TRUNC ranges never EXTRACT(MONTH) = number
- For comparing two months always use this exact pattern:
  SELECT
  ROUND(SUM(CASE WHEN order_date >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 months')
            AND order_date < DATE_TRUNC('month', CURRENT_DATE)
            THEN sales ELSE 0 END)::numeric, 2) AS last_month_sales,
  ROUND(SUM(CASE WHEN order_date >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
            AND order_date < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 months')
            THEN sales ELSE 0 END)::numeric, 2) AS month_before_sales
  FROM orders
- NEVER use BETWEEN for date comparisons always use >= and <
- NEVER use GROUP BY for comparison queries
- NEVER use GROUP BY column numbers like GROUP BY 1,2,3,4
- For month name: TO_CHAR(order_date, 'Month')
- For year-month grouping: TO_CHAR(order_date, 'YYYY-MM')
- Always use INTERVAL '1 months' not INTERVAL '1 month'
- Always cast string dates to ::date when using DATE_TRUNC
- Always alias all columns with meaningful names using AS
- Never add comments in SQL return only the raw query

LIMIT RULES:
- NEVER use LIMIT for breakdown questions: by category, by region, by segment, by state, by city, by year, by month, by quarter
- Only use LIMIT when user says: top N, bottom N, highest, lowest, most, least
- The word 'by' means show all groups never limit them

RETURNS TABLE RULES:
- For refund analysis always use SUM(refund_amount) from returns table
- For return reasons use GROUP BY reason from returns table
- Always use LEFT JOIN returns with orders when calculating return rates by group
- Always prefix columns with table alias when using JOIN to avoid ambiguity

{context_str}

Current question: {question}"""

    response = requests.post(
        url="http://192.168.15.220:11434/api/generate",
        headers={"Content-Type": "application/json"},
        json={
            "model": "gemma3:12b",
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.1,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
                "num_ctx": 8192
            }
        },
        timeout=120
    )
    sql = response.json()['response'].strip()
    sql = sql.replace("```sql", "").replace("```", "").strip()
    if sql.endswith(";"):
        sql = sql[:-1]
    return sql


def run_query(sql, question=""):
    try:
        sql = re.sub(r'ROUND\(([^,]+?)(?<!::numeric),\s*2\)', lambda m: f'ROUND(({m.group(1)})::numeric, 2)', sql, flags=re.IGNORECASE)
        sql = re.sub(r'ROUND\((.+?\* 100\))::numeric,\s*2\)', lambda m: f'ROUND(({m.group(1)})::numeric, 2)', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\bOrders\b', 'orders', sql)
        sql = re.sub(r'\bReturns\b', 'returns', sql)
        sql = re.sub(r"INTERVAL '(\d+) (month|year|day|week)'", lambda m: f"INTERVAL '{m.group(1)} {m.group(2)}s'", sql, flags=re.IGNORECASE)

        if re.search(r'EXTRACT\s*\(\s*MONTH\s+FROM', sql, re.IGNORECASE):
            if not re.search(r'EXTRACT\s*\(\s*YEAR\s+FROM', sql, re.IGNORECASE):
                if not re.search(r'DATE_TRUNC', sql, re.IGNORECASE):
                    current_year = datetime.now().year
                    sql = re.sub(r'\bWHERE\b', f'WHERE EXTRACT(YEAR FROM order_date) = {current_year} AND', sql, count=1, flags=re.IGNORECASE)
                    print(f"  🔧 Fixed: Added year filter {current_year}")

        def fix_extract_month(m):
            col = m.group(1)
            month_num = int(m.group(2))
            current_month = datetime.now().month
            months_ago = (current_month - month_num) % 12
            if months_ago == 0:
                months_ago = 12
            return (f"{col} >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{months_ago} months') "
                   f"AND {col} < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{months_ago - 1} months')")

        sql = re.sub(r"EXTRACT\s*\(\s*MONTH\s+FROM\s+(\w+)\s*\)\s*=\s*(\d+)", fix_extract_month, sql, flags=re.IGNORECASE)
        sql = re.sub(r"DATE_TRUNC\(('(?:quarter|month|year|week|day)'),\s*'(\d{4}-\d{2}-\d{2})'\)", lambda m: f"DATE_TRUNC({m.group(1)}, '{m.group(2)}'::date)", sql, flags=re.IGNORECASE)

        if any(word in question.lower() for word in ["cumulative", "running total", "accumulated"]):
            sql = re.sub(r'\bSUM\((\w+)\)\s+OVER\s*\(', r'SUM(SUM(\1)) OVER (', sql, flags=re.IGNORECASE)
            sql = re.sub(r'ROUND\(SUM\((\w+)\)\s+OVER\s*\(([^)]+)\)::numeric', r'ROUND(SUM(SUM(\1)) OVER (\2)::numeric', sql, flags=re.IGNORECASE)
            print(f"  🔧 Fixed: Corrected cumulative window function")

        if re.search(r'GROUP BY\s+\d+\s*,\s*\d+', sql, re.IGNORECASE):
            if re.search(r'SUM\s*\(', sql, re.IGNORECASE):
                sql = re.sub(r'GROUP BY[\s\d,]+', '', sql, flags=re.IGNORECASE)
                print(f"  🔧 Fixed: Removed illegal GROUP BY with aggregates")

        if re.search(r'\bJOIN\b', sql, re.IGNORECASE) and table_columns:
            group_by_match = re.search(r'GROUP BY\s+(\w+)', sql, re.IGNORECASE)
            if group_by_match:
                col = group_by_match.group(1).lower()
                tables_with_col = [t for t, cols in table_columns.items() if col in cols]
                if len(tables_with_col) > 1:
                    main_table_match = re.search(r'FROM\s+(\w+)\s+(\w+)', sql, re.IGNORECASE)
                    if main_table_match:
                        main_alias = main_table_match.group(2)
                        sql = re.sub(rf'\bSELECT\s+{col}\b', f'SELECT {main_alias}.{col}', sql, flags=re.IGNORECASE)
                        sql = re.sub(rf'\bGROUP BY\s+{col}\b', f'GROUP BY {main_alias}.{col}', sql, flags=re.IGNORECASE)
                        print(f"  🔧 Fixed ambiguous column: {col} → {main_alias}.{col}")

        if table_columns:
            sql_keywords = {'where', 'from', 'join', 'on', 'and', 'or', 'not', 'in', 'is', 'as', 'by', 'order', 'group', 'having', 'select', 'distinct', 'count', 'sum', 'avg', 'max', 'min', 'round', 'extract', 'case', 'when', 'then', 'else', 'end', 'null', 'true', 'false', 'limit', 'offset', 'union', 'all', 'inner', 'outer', 'left', 'right', 'full', 'cross', 'natural', 'with'}

            for alias_match in re.finditer(r'\b([a-zA-Z_]\w*)\.(\w+)', sql):
                alias = alias_match.group(1)
                col = alias_match.group(2).lower()
                if alias.lower() in sql_keywords:
                    continue
                from_match = re.search(rf'(?:FROM|JOIN)\s+(\w+)\s+{re.escape(alias)}\b', sql, re.IGNORECASE)
                if from_match:
                    table = from_match.group(1).lower()
                    if table in table_columns and col not in table_columns[table]:
                        for other_table, cols in table_columns.items():
                            if col in cols and other_table != table:
                                other_alias_match = re.search(rf'(?:FROM|JOIN)\s+{other_table}\s+([a-zA-Z_]\w*)\b', sql, re.IGNORECASE)
                                if other_alias_match:
                                    correct_alias = other_alias_match.group(1)
                                    sql = sql.replace(f'{alias}.{col}', f'{correct_alias}.{col}')
                                    print(f"  🔧 Fixed alias: {alias}.{col} → {correct_alias}.{col}")
                                    break

            for table_ref_match in re.finditer(r'\b(\w+)\.(\w+)\b', sql):
                ref_table = table_ref_match.group(1).lower()
                ref_col = table_ref_match.group(2).lower()
                if ref_table in sql_keywords:
                    continue
                if ref_table in table_columns and ref_col not in table_columns[ref_table]:
                    for other_table, cols in table_columns.items():
                        if ref_col in cols and other_table != ref_table:
                            sql = sql.replace(f'{ref_table}.{ref_col}', f'{other_table}.{ref_col}')
                            print(f"  🔧 Fixed table ref: {ref_table}.{ref_col} → {other_table}.{ref_col}")
                            break

        for normalized, quoted in special_columns.items():
            pattern = r'(?<!")(?<!\w)' + re.escape(normalized) + r'(?!\w)(?!")'
            sql = re.sub(pattern, quoted, sql, flags=re.IGNORECASE)

        sql = re.sub(r'ROUND\(SUM\((\w+)\)::numeric,\s*2\)(?!\s+AS)', lambda m: f'ROUND(SUM({m.group(1)})::numeric, 2) AS total_{m.group(1)}', sql, flags=re.IGNORECASE)

        breakdown_words = [" by ", "by category", "by region", "by segment", "by state", "by city", "by year", "by month", "by quarter", "by ship mode", "trends", "monthly", "quarterly"]
        limit_words = ["highest", "lowest", "most", "least", "top", "bottom", "best", "worst", "maximum", "minimum"]

        is_breakdown = any(word in question.lower() for word in breakdown_words)
        is_specific = any(word in question.lower() for word in limit_words)

        if is_breakdown and not is_specific:
            sql = re.sub(r'LIMIT\s+1', '', sql, flags=re.IGNORECASE)

        group_match = re.search(r'GROUP BY (\w+)', sql, re.IGNORECASE)
        select_match = re.search(r'SELECT (.+?) FROM', sql, re.IGNORECASE)
        if group_match and select_match:
            group_col = group_match.group(1)
            select_cols = select_match.group(1)
            if group_col.lower() not in select_cols.lower():
                sql = sql.replace('SELECT ', f'SELECT {group_col}, ', 1)

        with engine.connect() as conn:
            result = pd.read_sql(text(sql), conn)

            for col in result.columns:
                if 'discount' in col.lower():
                    if result[col].between(0, 1).all():
                        result[col] = (result[col] * 100).round(2)
                        result = result.rename(columns={col: col + '_%'})

            return result

    except Exception as e:
        print(f"  ❌ SQL Error: {e}")
        return None


def summarise(question, result):
    prompt = f"""You are an expert Sales Analytics Assistant for a retail company.
The user asked: '{question}'
The database returned: {result.to_string()}

Rules:
- Respond in strictly 2 sentences like a business analyst presenting findings
- Be insightful — mention what the number means for the business. Dont say much just 2/3 words of insight is enough.
- Example: 'West region leads with $725,457 in sales, contributing 32% of total revenue. This makes it the strongest market and worth prioritising for future investment.'
- Use exact names and numbers from the data
- Never add labels like 'row 0' or 'index 0' before names
- Never make up numbers not in the data
- Do not start with 'I' or 'According to'
- The leftmost column is the INDEX not the answer
- If result shows a single number use it directly
- If result has multiple rows present all clearly one by one
- If result has two columns that look like a comparison say: first period had X, second period had Y, difference is Z
- If discount value is between 0 and 1 multiply by 100 and show as percentage
- Never say 'round' as a column name always use the alias name"""

    response = requests.post(
        url="http://192.168.15.220:11434/api/generate",
        headers={"Content-Type": "application/json"},
        json={
            "model": "gemma3:12b",
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.4,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
                "num_ctx": 4096
            }
        },
        timeout=120
    )
    return response.json()['response'].strip()


def ask(question):
    global context
    print(f"\n  ┌─ Question: {question}")

    q_lower = question.lower().strip().rstrip("!?.")
    chat_exact = [
        "hello", "hi", "hey", "good morning", "good afternoon",
        "good evening", "how are you", "thanks", "thank you",
        "bye", "goodbye", "see you", "thats all", "that's all",
        "i don't understand", "i dont understand",
        "what can you do", "good job", "well done", "nice work"
    ]
    db_keywords = [
        "sales", "profit", "orders", "returns", "customers", "region",
        "category", "segment", "state", "city", "discount", "quantity",
        "ship", "delivery", "product", "refund", "revenue", "margin",
        "rate", "total", "average", "highest", "lowest", "top", "bottom",
        "best", "worst", "month", "year", "quarter", "week", "trend",
        "compare", "how many", "how much", "which", "what is", "show me",
        "give me", "list", "by how", "by much", "cumulative"
    ]
    has_db_keyword = any(kw in q_lower for kw in db_keywords)
    is_exact_chat = q_lower in chat_exact
    if is_exact_chat and not has_db_keyword:
        return chat_with_user(question)

    database_info_phrases = [
        "what is this database", "what database", "about this database",
        "what data", "what does this contain", "what can you answer",
        "what can you tell", "what information", "what is this about",
        "tell me about this database", "what do you know",
        "what is the name of the database", "name of the database",
        "what is the database name", "database name", "what is your database",
        "what are you", "what do you do", "how can you help",
        "what is the name of this database", "name of this database",
        "what is this", "which database"
    ]
    if any(phrase in question.lower() for phrase in database_info_phrases):
        return """This is the Superstore Sales Database containing real sales transactions from 2022 to 2026. Here is what it contains:

  📦  Orders     — 9,994 sales transactions with order details, shipping, customers, products and financials
  🔄  Returns    — 1,999 product returns with return reasons, quantities and refund amounts
  👤  Customers  — 793 unique customers across Consumer, Corporate and Home Office segments
  🌍  Geography  — 4 regions (East, West, Central, South) covering 49 US states
  📦  Products   — 3 categories: Furniture, Technology and Office Supplies
  💰  Financials — sales revenue, profit, quantity and discounts

  You can ask me anything about sales, profit, returns, customers, regions, products, delivery times, discounts and trends!"""

    export_phrases = ["download", "save", "export", "word doc", "word document", "save as doc", "generate report"]
    if any(phrase in question.lower() for phrase in export_phrases):
        if context:
            msg = save_to_word(context)
            return msg
        else:
            return "No conversation to save yet. Ask some questions first, then say download!"

    date_ctx = inject_date_into_prompt(question)
    sql = text_to_sql(question, date_ctx)
    print(f"  ├─ SQL: {sql}")

    if "UNRELATED" in sql.upper():
        if context and len(question.split()) <= 6:
            retry_q = f"{question} (referring to previous answer about {context[-1]['question']})"
            sql = text_to_sql(retry_q, date_ctx)
            if "UNRELATED" in sql.upper() or not sql.strip().upper().startswith(("SELECT", "WITH")):
                return chat_with_user(question)
        else:
            return chat_with_user(question)

    if not sql.strip().upper().startswith(("SELECT", "WITH")):
        return chat_with_user(question)

    result = run_query(sql, question)

    if result is None or result.empty:
        return "Couldn't find data for that. Try rephrasing — for example: 'total sales last month' or 'profit by region'."

    detail_phrases = ["give me details", "show me details", "tell me about", "all orders", "all data", "give me all"]
    if any(phrase in question.lower() for phrase in detail_phrases) and not needs_visualisation(question):
        print(f"\n  📊 Results — {len(result)} records found:\n")
        print(result.to_string(index=False))
        answer = f"Found {len(result)} records for your query."
        if result is not None and not result.empty:
            context.append({"question": question, "sql": sql, "answer": answer, "result": result.to_string()})
            if len(context) > 10:
                context.pop(0)
        return answer

    if needs_visualisation(question):
        create_visualisation(result, question)

    answer = summarise(question, result)

    if result is not None and not result.empty:
        context.append({
            "question": question,
            "sql": sql,
            "answer": answer,
            "result": result.to_string() if result is not None else ""
        })
        if len(context) > 10:
            context.pop(0)

    return answer


if __name__ == "__main__":
    print()
    print("  Welcome to the Superstore Sales Database Chatbot! 🤖")
    print()
    print("  Ask me anything about sales, returns, customers,")
    print("  regions, products, delivery and discounts.")
    print()
    print("  Commands:")
    print("  'exit'      — quit the chatbot")
    print("  'clear'     — reset conversation history")
    print("  'context'   — show conversation history")
    print("  'download'  — save chat as Word doc")
    print()
    print("=" * 60)

    while True:
        print()
        question = input("  You: ")

        if question.strip().lower() == "exit":
            print()
            print("=" * 60)
            print("  Thank you for using the Superstore Sales Chatbot!")
            print("  Goodbye! 👋")
            print("=" * 60)
            print()
            break

        if question.strip().lower() == "clear":
            context = []
            print("  ✅ Conversation history cleared!")
            continue

        if question.strip().lower() == "context":
            if not context:
                print("  No conversation history yet!")
            else:
                print(f"\n  📝 Last {len(context)} exchanges:")
                for i, exchange in enumerate(context):
                    print(f"\n  {i+1}. Q: {exchange['question']}")
                    print(f"     A: {exchange['answer']}")
            continue

        if question.strip() == "":
            print("  ⚠️  Please type a question!")
            continue

        print("  🤖 Thinking...")
        answer = ask(question)

        print(f"\n  └─ Answer:")
        wrapped = textwrap.fill(answer, width=70, initial_indent="     ", subsequent_indent="     ")
        print(wrapped)
        print()
        print("  " + "─" * 56)