"""
SQLite 数据库模块 — 替代 cache.json，支持全量工参存储和快速查询
"""
import sqlite3
import json
import os
import threading

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache.db")
db_lock = threading.Lock()

# 标准列（17个，按前端展示顺序）
STANDARD_COLS = [
    "技术制式", "运营商", "设备商", "基站名", "基站ID",
    "小区名", "PCI", "小区ID", "下行频点", "下倾角",
    "挂高", "方位角", "经度", "纬度", "共享", "频段", "TAC",
]

# 所有列（标准列 + 元数据列）
ALL_COLS = STANDARD_COLS + ["_文件名", "_工作表", "_raw"]


def get_conn():
    """获取数据库连接（同一线程复用）"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db():
    """初始化数据库表结构"""
    null_count = 0
    with db_lock:
        conn = get_conn()
        try:
            # 记录表
            cols_def = ",\n            ".join(
                [f'"{c}" TEXT' for c in STANDARD_COLS] +
                ['"_文件名" TEXT', '"_工作表" TEXT', '"_raw" TEXT', '"_search" TEXT']
            )
            conn.execute(f"""
                CREATE TABLE IF NOT EXISTS records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    {cols_def}
                )
            """)

            # 兼容旧表：添加 _search 列
            try:
                conn.execute('ALTER TABLE records ADD COLUMN "_search" TEXT')
            except Exception:
                pass

            # 索引
            for col in ["小区名", "基站名", "PCI", "小区ID", "基站ID", "TAC", "下行频点", "_文件名", "技术制式", "运营商", "_search"]:
                try:
                    conn.execute(f'CREATE INDEX IF NOT EXISTS idx_{col} ON records("{col}")')
                except Exception:
                    pass

            # 小区ID 规范化迁移：长格式(基站ID+小区标识)→截短标识
            try:
                conn.execute('''
                    UPDATE records SET "小区ID" = substr("小区ID", length("基站ID")+1)
                    WHERE length("基站ID") > 0 AND "小区ID" LIKE "基站ID" || '%'
                      AND length("小区ID") > length("基站ID")
                ''')
                conn.commit()
            except Exception:
                pass

            # 文件来源表
            conn.execute("""
                CREATE TABLE IF NOT EXISTS file_sources (
                    filename TEXT PRIMARY KEY,
                    carrier TEXT,
                    tech TEXT,
                    count INTEGER,
                    sheets TEXT
                )
            """)

            # 检查是否需要补全 _search
            null_count = conn.execute('SELECT COUNT(*) as cnt FROM records WHERE "_search" IS NULL').fetchone()["cnt"]
            conn.commit()
        finally:
            conn.close()

    # 在锁外补全 _search（可能耗时较长）
    if null_count > 0:
        _backfill_search()


def clear_all():
    """清空所有数据"""
    with db_lock:
        conn = get_conn()
        try:
            conn.execute("DELETE FROM records")
            conn.execute("DELETE FROM file_sources")
            conn.commit()
        finally:
            conn.close()


def _make_search(rec: dict) -> str:
    """生成搜索辅助字段：合并关键列，小写去连字符"""
    parts = []
    for key in ["小区名", "PCI", "小区ID", "基站名"]:
        v = str(rec.get(key, "") or "")
        parts.append(v.lower().replace("-", ""))
    return " ".join(parts)


def _backfill_search():
    """补全已有数据的 _search 字段"""
    conn = get_conn()
    try:
        batch_size = 5000
        while True:
            rows = conn.execute(
                'SELECT id, "小区名","PCI","小区ID","基站名" FROM records WHERE "_search" IS NULL LIMIT ?',
                (batch_size,)
            ).fetchall()
            if not rows:
                break
            for r in rows:
                search_val = _make_search({
                    "小区名": r["小区名"], "PCI": r["PCI"],
                    "小区ID": r["小区ID"], "基站名": r["基站名"]
                })
                conn.execute('UPDATE records SET "_search" = ? WHERE id = ?',
                             (search_val, r["id"]))
            conn.commit()
            print(f"[搜索优化] 已补全 {len(rows)} 条 _search 字段")
    finally:
        conn.close()


def insert_records(recs: list):
    """批量插入记录
    recs: [dict], 每个 dict 包含标准列 + _文件名/_工作表/_raw
    """
    if not recs:
        return 0
    with db_lock:
        conn = get_conn()
        try:
            all_cols = STANDARD_COLS + ["_文件名", "_工作表", "_raw", "_search"]
            placeholders = ", ".join(["?" for _ in all_cols])
            col_names = ", ".join([f'"{c}"' for c in all_cols])
            sql = f"INSERT INTO records ({col_names}) VALUES ({placeholders})"

            rows = []
            for r in recs:
                row = []
                for c in all_cols:
                    if c == "_raw" and isinstance(r.get(c), dict):
                        val = json.dumps(r.get(c), ensure_ascii=False)
                    elif c == "_search":
                        val = _make_search(r)
                    else:
                        val = str(r.get(c, "") or "")
                    row.append(val)
                rows.append(tuple(row))

            conn.executemany(sql, rows)
            conn.commit()
            return len(rows)
        finally:
            conn.close()


def get_total_count():
    """获取总记录数"""
    conn = get_conn()
    try:
        row = conn.execute("SELECT COUNT(*) as cnt FROM records").fetchone()
        return row["cnt"] if row else 0
    finally:
        conn.close()


def get_unique_station_count():
    """获取唯一基站数"""
    conn = get_conn()
    try:
        row = conn.execute(
            "SELECT COUNT(DISTINCT 基站名 || '|' || COALESCE(经度,'') || '|' || COALESCE(纬度,'')) as cnt FROM records"
        ).fetchone()
        return row["cnt"] if row else 0
    finally:
        conn.close()


# 多字段搜索的字段映射：前端参数名 → 数据库列名
FILTER_FIELD_MAP = {
    "制式": "技术制式",
    "TAC": "TAC",
    "基站ID": "基站ID",
    "小区ID": "小区ID",
    "PCI": "PCI",
    "频点": "下行频点",
    "基站名": "基站名",
    "小区名": "小区名",
}

# 数字型字段（优先精确匹配 =，走索引）
NUMERIC_FIELDS = {"基站ID", "小区ID", "PCI", "TAC", "频点"}


def search_records(query: str = "", page: int = 1, per_page: int = 50, filters: dict = None):
    """搜索记录，返回 (records_list, total_count)
    
    支持两种模式：
    1. 传统模式：query 非空时在 _search 列中模糊匹配
    2. 多字段模式：filters 指定各字段值，AND 组合，智能匹配（数字=精确，文本=模糊）
    """
    q = str(query).strip() if query else ""
    conn = get_conn()
    try:
        conditions = []
        params = []

        # 多字段过滤（AND 组合）
        if filters:
            for param_name, db_col in FILTER_FIELD_MAP.items():
                val = str(filters.get(param_name, "")).strip()
                if not val:
                    continue
                # 制式特殊处理
                if param_name == "制式":
                    if val.upper() in ("4G", "LTE"):
                        conditions.append(f'"{db_col}" LIKE ?')
                        params.append("%4G%")
                    elif val.upper() in ("5G", "NR"):
                        conditions.append(f'"{db_col}" LIKE ?')
                        params.append("%5G%")
                    else:
                        conditions.append(f'"{db_col}" LIKE ?')
                        params.append(f"%{val}%")
                # 数字型字段：纯数字 → 精确匹配 =（走索引）；含非数字字符 → LIKE 回退
                elif param_name in NUMERIC_FIELDS:
                    if val.isdigit():
                        conditions.append(f'"{db_col}" = ?')
                        params.append(val)
                    else:
                        conditions.append(f'"{db_col}" LIKE ?')
                        params.append(f"%{val}%")
                # 文本型字段：模糊 LIKE
                else:
                    conditions.append(f'"{db_col}" LIKE ?')
                    params.append(f"%{val}%")

        # 传统模糊搜索
        if q:
            q_lower = q.lower().replace("-", "")
            q_parts = [p for p in q_lower.split() if p]
            or_conds = []
            or_params = []
            or_conds.append('"_search" LIKE ?')
            or_params.append(f"%{q_lower}%")
            for part in q_parts:
                or_conds.append('"_search" LIKE ?')
                or_params.append(f"%{part}%")
            conditions.append("(" + " OR ".join(or_conds) + ")")
            params.extend(or_params)

        # 无任何条件 → 返回全部
        cols = ", ".join([f'"{c}"' for c in STANDARD_COLS] + ['"_文件名"', '"_工作表"', 'id'])
        if not conditions:
            total = conn.execute("SELECT COUNT(*) as cnt FROM records").fetchone()["cnt"]
            offset = (page - 1) * per_page
            rows = conn.execute(
                f"SELECT {cols} FROM records ORDER BY id LIMIT ? OFFSET ?",
                (per_page, offset)
            ).fetchall()
            return [dict(r) for r in rows], total

        where = " AND ".join(conditions)
        sql = f"SELECT {cols} FROM records WHERE {where}"
        total = conn.execute(f"SELECT COUNT(*) as cnt FROM records WHERE {where}", params).fetchone()["cnt"]

        offset = (page - 1) * per_page
        rows = conn.execute(
            f"{sql} ORDER BY id LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()

        return [dict(r) for r in rows], total
    finally:
        conn.close()


def get_all_records(max_rows: int = 999999):
    """获取全部记录（用于导出）"""
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT * FROM records ORDER BY id LIMIT ?", (max_rows,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_filtered_records(tech_keyword: str, max_rows: int = 999999):
    """按技术制式过滤"""
    conn = get_conn()
    try:
        rows = conn.execute(
            'SELECT * FROM records WHERE "技术制式" LIKE ? ORDER BY id LIMIT ?',
            (f"%{tech_keyword}%", max_rows)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def delete_by_filename(filename: str):
    """按文件名删除记录"""
    with db_lock:
        conn = get_conn()
        try:
            conn.execute('DELETE FROM records WHERE "_文件名" = ?', (filename,))
            conn.commit()
        finally:
            conn.close()


def delete_by_filename_and_sheets(filename: str, sheets: list):
    """按文件名+sheet列表删除记录"""
    with db_lock:
        conn = get_conn()
        try:
            placeholders = ", ".join(["?" for _ in sheets])
            conn.execute(
                f'DELETE FROM records WHERE "_文件名" = ? AND "_工作表" IN ({placeholders})',
                [filename] + sheets
            )
            conn.commit()
        finally:
            conn.close()


# ============ file_sources 操作 ============

def get_file_sources():
    """获取所有文件来源"""
    conn = get_conn()
    try:
        rows = conn.execute("SELECT * FROM file_sources").fetchall()
        result = {}
        for r in rows:
            result[r["filename"]] = {
                "carrier": r["carrier"],
                "tech": r["tech"],
                "count": r["count"],
                "sheets": json.loads(r["sheets"]) if r["sheets"] else [],
            }
        return result
    finally:
        conn.close()


def upsert_file_source(filename: str, carrier: str, tech: str, count: int, sheets: list):
    """插入或更新文件来源"""
    with db_lock:
        conn = get_conn()
        try:
            sheets_json = json.dumps(sheets, ensure_ascii=False)
            conn.execute("""
                INSERT INTO file_sources (filename, carrier, tech, count, sheets)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(filename) DO UPDATE SET
                    carrier=excluded.carrier,
                    tech=excluded.tech,
                    count=excluded.count,
                    sheets=excluded.sheets
            """, (filename, carrier, tech, count, sheets_json))
            conn.commit()
        finally:
            conn.close()


def remove_file_source(filename: str):
    """删除文件来源"""
    with db_lock:
        conn = get_conn()
        try:
            conn.execute("DELETE FROM file_sources WHERE filename = ?", (filename,))
            conn.commit()
        finally:
            conn.close()


def get_stats():
    """获取统计信息"""
    conn = get_conn()
    try:
        total = conn.execute("SELECT COUNT(*) as cnt FROM records").fetchone()["cnt"]
        stations = conn.execute(
            "SELECT COUNT(DISTINCT 基站名 || '|' || COALESCE(经度,'') || '|' || COALESCE(纬度,'')) as cnt FROM records"
        ).fetchone()["cnt"]
        files = conn.execute("SELECT COUNT(*) as cnt FROM file_sources").fetchone()["cnt"]
        return {"total": total, "stations": stations, "files": files}
    finally:
        conn.close()
