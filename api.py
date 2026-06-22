# ============================================================
# api.py — FastAPI backend for Superstore Sales Chatbot
# Run with: uvicorn api:app --host 0.0.0.0 --port 8000 --reload
# ============================================================

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response, HTMLResponse
from pydantic import BaseModel
from typing import Optional
import sys
import os
import io
import re
import json
import uuid
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import chatbot_code1 as chatbot

app = FastAPI(title="Superstore Sales Chatbot API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── SESSION STORAGE (MEMORY ONLY — cleared on server restart) ──
sessions = {}           # session_id → {id, title, created_at, messages}
session_contexts = {}   # session_id → {exchanges: [...], summary: "..."}

def load_sessions():
    return sessions

def save_sessions(data):
    global sessions
    sessions = data

# Track active session
active_session_id = None

MAX_RAW_EXCHANGES = 3   # keep last 3 raw exchanges per session

class QuestionRequest(BaseModel):
    question: str
    session_id: Optional[str] = None

class AnswerResponse(BaseModel):
    answer: str
    has_chart: bool = False
    session_id: str = ""

class TableRequest(BaseModel):
    table_name: str
    limit: int = 100
    offset: int = 0

class SessionRequest(BaseModel):
    session_id: str


@app.get("/")
def root():
    """Serve the main dashboard."""
    import os
    html_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    if os.path.exists(html_file):
        with open(html_file, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return {"status": "Superstore Chatbot API is running"}


# ── SESSION ENDPOINTS ─────────────────────────────────────────

@app.post("/session/new")
def new_session():
    """Create a new chat session."""
    global active_session_id
    session_id = str(uuid.uuid4())
    sessions_data = load_sessions()
    sessions_data[session_id] = {
        "id": session_id,
        "title": "New Chat",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "messages": []
    }
    save_sessions(sessions_data)
    # Init a fresh context bucket for this session
    session_contexts[session_id] = {"exchanges": [], "summary": ""}
    # Load empty context into chatbot
    chatbot.context.clear()
    chatbot.latest_chart_buf = None
    active_session_id = session_id
    return {"session_id": session_id}


@app.get("/sessions")
def get_sessions():
    """Return last 5 sessions sorted by newest first."""
    sessions = load_sessions()
    session_list = sorted(
        sessions.values(),
        key=lambda x: x["created_at"],
        reverse=True
    )[:5]
    return {"sessions": session_list}


@app.post("/session/load")
def load_session(req: SessionRequest):
    """Load a previous session and restore its context fully."""
    global active_session_id
    sessions = load_sessions()
    if req.session_id not in sessions:
        return {"error": "Session not found"}
    session = sessions[req.session_id]

    # Restore chatbot.context from this session's context bucket
    chatbot.context.clear()
    chatbot.latest_chart_buf = None
    if req.session_id in session_contexts:
        ctx = session_contexts[req.session_id]
        for ex in ctx["exchanges"]:
            chatbot.context.append(ex)
    # If session_contexts lost this session (e.g. partial restart), init empty
    else:
        session_contexts[req.session_id] = {"exchanges": [], "summary": ""}

    active_session_id = req.session_id
    return {"session": session}


@app.delete("/session/{session_id}")
def delete_session(session_id: str):
    """Delete a session."""
    sessions = load_sessions()
    if session_id in sessions:
        del sessions[session_id]
        save_sessions(sessions)
    return {"status": "deleted"}


@app.get("/health")
def health():
    try:
        from sqlalchemy import text
        with chatbot.engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        db_ok = True
    except Exception:
        db_ok = False

    try:
        import requests
        r = requests.post(
            url="http://192.168.15.220:11434/api/generate",
            json={"model": "gemma3:12b", "prompt": "hi", "stream": False},
            timeout=60
        )
        ai_ok = r.status_code == 200
    except Exception:
        ai_ok = False

    return {
        "database": "connected" if db_ok else "disconnected",
        "ai_model": "connected" if ai_ok else "disconnected",
        "status": "ok" if (db_ok and ai_ok) else "degraded"
    }


def _load_session_context(session_id: str):
    """Load this session's context into chatbot.context (summary + last 3 exchanges)."""
    chatbot.context.clear()
    ctx = session_contexts.get(session_id)
    if not ctx:
        return
    # Inject summary as a synthetic first exchange so Gemma3 sees it
    summary = ctx.get("summary", "")
    if summary:
        chatbot.context.append({
            "question": "__summary__",
            "sql": "",
            "answer": summary,
            "result": ""
        })
    # Then inject last MAX_RAW_EXCHANGES raw exchanges
    for ex in ctx["exchanges"][-MAX_RAW_EXCHANGES:]:
        chatbot.context.append(ex)


@app.post("/ask", response_model=AnswerResponse)
def ask_question(req: QuestionRequest):
    global active_session_id

    # ── Determine session ─────────────────────────────────────
    incoming_sid = req.session_id

    if incoming_sid and incoming_sid in session_contexts:
        # Known session — use it
        session_id = incoming_sid
        if session_id != active_session_id:
            # Switching to a different session — restore its context
            _load_session_context(session_id)
            active_session_id = session_id
        else:
            # Same session already active — just make sure context is loaded
            if not chatbot.context:
                _load_session_context(session_id)
    else:
        # No session or unknown — create a new one
        session_id = str(uuid.uuid4())
        sessions_data = load_sessions()
        sessions_data[session_id] = {
            "id": session_id,
            "title": "New Chat",
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "messages": []
        }
        save_sessions(sessions_data)
        session_contexts[session_id] = {"exchanges": [], "summary": ""}
        chatbot.context.clear()
        chatbot.latest_chart_buf = None
        active_session_id = session_id

    # ── Ask chatbot ───────────────────────────────────────────
    chatbot.latest_chart_buf = None
    answer = chatbot.ask(req.question)
    has_chart = chatbot.latest_chart_buf is not None

    # ── Save new exchange into session context ────────────────
    if chatbot.context:
        real_exchanges = [ex for ex in chatbot.context if ex.get("question") != "__summary__"]
        session_contexts[session_id]["exchanges"] = real_exchanges

    # ── Save messages to session store ────────────────────────
    sessions_data = load_sessions()
    if session_id in sessions_data:
        session = sessions_data[session_id]
        if session["title"] == "New Chat" and req.question:
            session["title"] = req.question[:50]
        session["messages"].append({"role": "user", "content": req.question})
        session["messages"].append({"role": "bot", "content": answer})
        save_sessions(sessions_data)

    return AnswerResponse(answer=answer, has_chart=has_chart, session_id=session_id)


@app.get("/chart/latest")
def get_latest_chart():
    """Serve the latest chart from memory — no file saved."""
    if chatbot.latest_chart_buf is None:
        return Response(status_code=404)
    chatbot.latest_chart_buf.seek(0)
    return Response(
        content=chatbot.latest_chart_buf.read(),
        media_type="image/png"
    )

@app.get("/chart-data")
def get_chart_data():
    """Return last query result as JSON for interactive Chart.js rendering."""
    if not chatbot.context:
        return Response(status_code=404)
    
    last = chatbot.context[-1]
    result_str = last.get('result', '')
    question = last.get('question', '').lower()
    
    if not result_str:
        return Response(status_code=404)
    
    try:
        import pandas as pd
        import io
        df = pd.read_csv(io.StringIO(result_str), sep=r'\s+', engine='python')
        df = df.dropna()
        if len(df.columns) < 2:
            return Response(status_code=404)

        is_line = any(w in question for w in [
            'trend', 'monthly', 'quarterly', 'yearly',
            'by month', 'by year', 'by quarter', 'over time', 'cumulative'
        ])

        labels = df.iloc[:, 0].astype(str).tolist()
        values = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0).tolist()

        return {
            "type": "line" if is_line else "bar",
            "labels": labels,
            "values": values,
            "title": last.get('question', '')
        }
    except Exception as e:
        print(f"  ⚠️ chart-data error: {e}")
        return Response(status_code=404)

@app.post("/clear")
def clear_context():
    global active_session_id
    chatbot.context.clear()
    chatbot.latest_chart_buf = None
    if active_session_id and active_session_id in session_contexts:
        session_contexts[active_session_id] = {"exchanges": [], "summary": ""}
    return {"status": "cleared"}


@app.get("/context")
def get_context():
    return {
        "count": len(chatbot.context),
        "history": [
            {
                "question": ex["question"],
                "answer": ex["answer"].split('\n\n  💡')[0].strip()
            }
            for ex in chatbot.context
        ]
    }


@app.get("/stats")
def get_stats():
    try:
        from sqlalchemy import text
        with chatbot.engine.connect() as conn:
            total_sales = conn.execute(text("SELECT ROUND(SUM(sales)::numeric, 2) FROM orders")).fetchone()[0]
            total_orders = conn.execute(text("SELECT COUNT(DISTINCT order_id) FROM orders")).fetchone()[0]
            total_profit = conn.execute(text("SELECT ROUND(SUM(profit)::numeric, 2) FROM orders")).fetchone()[0]
            total_customers = conn.execute(text("SELECT COUNT(DISTINCT customer_name) FROM orders")).fetchone()[0]
            return_rate = conn.execute(text("""
                SELECT ROUND((COUNT(DISTINCT r.return_id)::numeric /
                NULLIF(COUNT(DISTINCT o.order_id)::numeric, 0)) * 100, 1)
                FROM orders o LEFT JOIN returns r ON r.order_id = o.order_id
            """)).fetchone()[0]

        return {
            "total_sales": float(total_sales or 0),
            "total_orders": int(total_orders or 0),
            "total_profit": float(total_profit or 0),
            "total_customers": int(total_customers or 0),
            "return_rate": float(return_rate or 0)
        }
    except Exception as e:
        return {"error": str(e)}


@app.post("/table")
def get_table(req: TableRequest):
    allowed_tables = ["orders", "returns","customers"]
    if req.table_name not in allowed_tables:
        return {"error": "Invalid table name"}
    try:
        from sqlalchemy import text
        import pandas as pd
        with chatbot.engine.connect() as conn:
            count = conn.execute(text(f"SELECT COUNT(*) FROM {req.table_name}")).fetchone()[0]
            df = pd.read_sql(
                text(f"SELECT * FROM {req.table_name} LIMIT {req.limit} OFFSET {req.offset}"),
                conn
            )
        return {
            "table": req.table_name,
            "total_rows": count,
            "offset": req.offset,
            "limit": req.limit,
            "columns": list(df.columns),
            "rows": df.fillna("").astype(str).values.tolist()
        }
    except Exception as e:
        return {"error": str(e)}


@app.post("/download")
def download_report():
    if not chatbot.context:
        return {"error": "No conversation to export yet"}

    result = chatbot.save_to_word(chatbot.context)

    if "chat_report_" in result:
        match = re.search(r'chat_report_\d+\.docx', result)
        if match:
            filename = match.group(0)
            if os.path.exists(filename):
                return FileResponse(
                    filename,
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    filename=filename
                )

    return {"message": result}