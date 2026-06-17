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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import chatbot_code1 as chatbot

app = FastAPI(title="Superstore Sales Chatbot API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class QuestionRequest(BaseModel):
    question: str

class AnswerResponse(BaseModel):
    answer: str
    has_chart: bool = False

class TableRequest(BaseModel):
    table_name: str
    limit: int = 100
    offset: int = 0


@app.get("/")
def root():
    """Serve the main dashboard."""
    import os
    html_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    if os.path.exists(html_file):
        with open(html_file, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return {"status": "Superstore Chatbot API is running"}


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
            timeout=10
        )
        ai_ok = r.status_code == 200
    except Exception:
        ai_ok = False

    return {
        "database": "connected" if db_ok else "disconnected",
        "ai_model": "connected" if ai_ok else "disconnected",
        "status": "ok" if (db_ok and ai_ok) else "degraded"
    }


@app.post("/ask", response_model=AnswerResponse)
def ask_question(req: QuestionRequest):
    # Reset chart buffer before each question
    chatbot.latest_chart_buf = None

    answer = chatbot.ask(req.question)

    has_chart = chatbot.latest_chart_buf is not None

    return AnswerResponse(answer=answer, has_chart=has_chart)


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


@app.post("/clear")
def clear_context():
    chatbot.context.clear()
    chatbot.latest_chart_buf = None
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
