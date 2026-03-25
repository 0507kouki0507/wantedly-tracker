"""全履歴を保存するSQLiteモジュール"""

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

DB_PATH = Path(__file__).parent / "data" / "analytics.db"


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: Path = DB_PATH) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with get_connection(db_path) as conn:
        conn.executescript("""
            -- 全スクレイプ履歴
            CREATE TABLE IF NOT EXISTS history (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                scraped_at  TEXT NOT NULL,
                article_type TEXT NOT NULL,   -- 募集 / ストーリー
                title       TEXT NOT NULL,
                status      TEXT,             -- 募集中 / 募集停止中
                pv          INTEGER DEFAULT 0,
                bookmark    INTEGER DEFAULT 0,
                oubo        INTEGER DEFAULT 0,
                ouen        INTEGER DEFAULT 0,
                likes       INTEGER DEFAULT 0
            );

            -- 日別差分計算用（前回値）
            CREATE TABLE IF NOT EXISTS last_values (
                article_key  TEXT PRIMARY KEY,
                pv           INTEGER DEFAULT 0,
                bookmark     INTEGER DEFAULT 0,
                oubo         INTEGER DEFAULT 0,
                ouen         INTEGER DEFAULT 0,
                likes        INTEGER DEFAULT 0,
                scraped_at   TEXT
            );
        """)


def save_history(conn: sqlite3.Connection, article_type: str, title: str,
                 status: str, pv: int, bookmark: int,
                 oubo: int, ouen: int, likes: int) -> None:
    now = datetime.now().isoformat()
    conn.execute("""
        INSERT INTO history (scraped_at, article_type, title, status,
                             pv, bookmark, oubo, ouen, likes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (now, article_type, title, status, pv, bookmark, oubo, ouen, likes))
    conn.commit()


def get_last(conn: sqlite3.Connection, key: str) -> Optional[sqlite3.Row]:
    return conn.execute(
        "SELECT * FROM last_values WHERE article_key = ?", (key,)
    ).fetchone()


def save_last(conn: sqlite3.Connection, key: str, pv: int, bookmark: int,
              oubo: int, ouen: int, likes: int) -> None:
    now = datetime.now().isoformat()
    conn.execute("""
        INSERT INTO last_values (article_key, pv, bookmark, oubo, ouen, likes, scraped_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(article_key) DO UPDATE SET
            pv=excluded.pv, bookmark=excluded.bookmark,
            oubo=excluded.oubo, ouen=excluded.ouen,
            likes=excluded.likes, scraped_at=excluded.scraped_at
    """, (key, pv, bookmark, oubo, ouen, likes, now))
    conn.commit()


def get_article_history(conn: sqlite3.Connection, title: str) -> pd.DataFrame:
    rows = conn.execute("""
        SELECT scraped_at, pv, bookmark, oubo, ouen, likes, status
        FROM history WHERE title = ?
        ORDER BY scraped_at ASC
    """, (title,)).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def get_all_history_by_day(conn: sqlite3.Connection) -> pd.DataFrame:
    """日付×記事の累計値テーブルを返す（日別差分計算用）"""
    rows = conn.execute("""
        SELECT date(scraped_at) AS date, article_type, title, status,
               MAX(pv) AS pv, MAX(oubo) AS oubo, MAX(likes) AS likes
        FROM history
        GROUP BY date(scraped_at), title
        ORDER BY title, date ASC
    """).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def get_all_latest(conn: sqlite3.Connection) -> pd.DataFrame:
    """各記事の最新スナップショットを返す"""
    rows = conn.execute("""
        SELECT h.article_type, h.title, h.status, h.pv, h.bookmark, h.oubo, h.ouen, h.likes,
               MIN(h.scraped_at) AS first_seen,
               MAX(h.scraped_at) AS last_seen
        FROM history h
        INNER JOIN (
            SELECT title, MAX(scraped_at) AS max_at
            FROM history GROUP BY title
        ) latest ON h.title = latest.title AND h.scraped_at = latest.max_at
        GROUP BY h.title
        ORDER BY h.pv DESC
    """).fetchall()
    return pd.DataFrame([dict(r) for r in rows])
