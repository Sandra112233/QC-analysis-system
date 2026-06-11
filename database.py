# database.py
# -*- coding: utf-8 -*-
import sqlite3
import json
from datetime import datetime, timedelta

DB_PATH = "qc_system.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_time TEXT NOT NULL,
            product_name TEXT,
            batch_no TEXT,
            spec TEXT,
            inspector TEXT,
            inspection_date TEXT,
            data_json TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def save_record(record):
    conn = get_connection()
    conn.execute("""
        INSERT INTO records (upload_time, product_name, batch_no, spec, inspector, inspection_date, data_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        record["upload_time"],
        record.get("product_name", ""),
        record.get("batch_no", ""),
        record.get("spec", ""),
        record.get("inspector", ""),
        record.get("inspection_date", ""),
        json.dumps(record["data"], ensure_ascii=False)
    ))
    conn.commit()
    conn.close()


def load_all_records():
    conn = get_connection()
    rows = conn.execute("SELECT * FROM records ORDER BY id DESC").fetchall()
    conn.close()
    records = []
    for row in rows:
        records.append({
            "id": row["id"],
            "upload_time": row["upload_time"],
            "product_name": row["product_name"],
            "batch_no": row["batch_no"],
            "spec": row["spec"],
            "inspector": row["inspector"],
            "inspection_date": row["inspection_date"],
            "data": json.loads(row["data_json"])
        })
    return records


def delete_records(ids):
    conn = get_connection()
    conn.executemany("DELETE FROM records WHERE id = ?", [(i,) for i in ids])
    conn.commit()
    conn.close()