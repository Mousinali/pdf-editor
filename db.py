"""MySQL storage for PDF editor session metadata (PDF bytes stay on disk).

Configure with env vars: MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DATABASE.
"""
import json
import os
from contextlib import contextmanager

import pymysql
from pymysql.cursors import DictCursor

CFG = {
    'host': os.environ.get('MYSQL_HOST', 'localhost'),
    'port': int(os.environ.get('MYSQL_PORT', '3306')),
    'user': os.environ.get('MYSQL_USER', 'root'),
    'password': os.environ.get('MYSQL_PASSWORD', ''),
}
DATABASE = os.environ.get('MYSQL_DATABASE', 'document_converter')

SCHEMAS = ["""
CREATE TABLE IF NOT EXISTS editor_sessions (
    sid        CHAR(32)     NOT NULL PRIMARY KEY,
    filename   VARCHAR(255) NOT NULL DEFAULT '',
    page_count INT UNSIGNED NOT NULL,
    size_bytes BIGINT UNSIGNED NOT NULL DEFAULT 0,
    created_at TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    KEY idx_updated_at (updated_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
""", """
CREATE TABLE IF NOT EXISTS operations (
    id           BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
    tool         VARCHAR(32)  NOT NULL,
    filename     VARCHAR(255) NOT NULL DEFAULT '',
    input_bytes  BIGINT UNSIGNED NOT NULL DEFAULT 0,
    output_bytes BIGINT UNSIGNED NULL,
    params       JSON NULL,
    status_code  SMALLINT UNSIGNED NOT NULL,
    success      TINYINT(1)   NOT NULL,
    error        VARCHAR(500) NULL,
    duration_ms  INT UNSIGNED NOT NULL,
    client_ip    VARCHAR(45)  NOT NULL DEFAULT '',
    created_at   TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_tool_created (tool, created_at),
    KEY idx_created (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
""", """
CREATE TABLE IF NOT EXISTS editor_saves (
    id           BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
    sid          CHAR(32)    NOT NULL,
    format       VARCHAR(8)  NOT NULL,
    edit_count   INT UNSIGNED NOT NULL,
    pages_kept   INT UNSIGNED NULL,
    output_bytes BIGINT UNSIGNED NOT NULL,
    created_at   TIMESTAMP   NOT NULL DEFAULT CURRENT_TIMESTAMP,
    KEY idx_sid (sid)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""]


def _connect(database=DATABASE):
    return pymysql.connect(**CFG, database=database, charset='utf8mb4',
                           cursorclass=DictCursor, autocommit=True, connect_timeout=5)


@contextmanager
def cursor():
    conn = _connect()
    try:
        with conn.cursor() as cur:
            yield cur
    finally:
        conn.close()


def init_db():
    """Create the database and tables if missing. Call once at startup."""
    conn = _connect(database=None)
    try:
        with conn.cursor() as cur:
            cur.execute(f'CREATE DATABASE IF NOT EXISTS `{DATABASE}` CHARACTER SET utf8mb4')
    finally:
        conn.close()
    with cursor() as cur:
        for ddl in SCHEMAS:
            cur.execute(ddl)


def create_session(sid, filename, page_count, size_bytes):
    with cursor() as cur:
        cur.execute('INSERT INTO editor_sessions (sid, filename, page_count, size_bytes) '
                    'VALUES (%s, %s, %s, %s)', (sid, filename[:255], page_count, size_bytes))


def session_exists(sid):
    with cursor() as cur:
        cur.execute('SELECT 1 FROM editor_sessions WHERE sid = %s', (sid,))
        return cur.fetchone() is not None


def touch_session(sid, size_bytes=None):
    with cursor() as cur:
        if size_bytes is None:
            cur.execute('UPDATE editor_sessions SET updated_at = CURRENT_TIMESTAMP WHERE sid = %s', (sid,))
        else:
            cur.execute('UPDATE editor_sessions SET size_bytes = %s, updated_at = CURRENT_TIMESTAMP '
                        'WHERE sid = %s', (size_bytes, sid))


def pop_expired(max_age_seconds):
    """Delete and return sids of sessions idle longer than max_age_seconds."""
    with cursor() as cur:
        cur.execute('SELECT sid FROM editor_sessions WHERE updated_at < NOW() - INTERVAL %s SECOND',
                    (int(max_age_seconds),))
        sids = [r['sid'] for r in cur.fetchall()]
        if sids:
            cur.execute('DELETE FROM editor_sessions WHERE sid IN (%s)' % ','.join(['%s'] * len(sids)), sids)
        return sids


def log_operation(tool, filename, input_bytes, output_bytes, params, status_code, error, duration_ms, client_ip):
    with cursor() as cur:
        cur.execute(
            'INSERT INTO operations (tool, filename, input_bytes, output_bytes, params, status_code, success, '
            'error, duration_ms, client_ip) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
            (tool, filename[:255], input_bytes, output_bytes, json.dumps(params) if params else None,
             status_code, int(status_code < 400), error[:500] if error else None, duration_ms, client_ip[:45]))


def log_editor_save(sid, fmt, edit_count, pages_kept, output_bytes):
    with cursor() as cur:
        cur.execute('INSERT INTO editor_saves (sid, format, edit_count, pages_kept, output_bytes) '
                    'VALUES (%s,%s,%s,%s,%s)', (sid, fmt, edit_count, pages_kept, output_bytes))
