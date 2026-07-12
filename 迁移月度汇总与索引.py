import json
import sqlite3
import time
from datetime import datetime
from pathlib import Path

from 数据库 import DB_PATH, DATA_DIR, init_db
from 分析引擎 import ensure_monthly_summary, month_int


def timed_query(conn, sql, params=(), repeats=3):
    durations = []
    row_count = 0
    for _ in range(repeats):
        start = time.perf_counter()
        rows = conn.execute(sql, params).fetchall()
        durations.append((time.perf_counter() - start) * 1000)
        row_count = len(rows)
    return {"ms": round(sum(durations) / len(durations), 2), "rows": row_count}


def backup_database():
    backup_dir = DATA_DIR / "备份"
    backup_dir.mkdir(parents=True, exist_ok=True)
    target = backup_dir / f"gmv诊断系统-月度汇总迁移前-{datetime.now().strftime('%Y%m%d-%H%M%S')}.sqlite3"
    source = sqlite3.connect(DB_PATH)
    dest = sqlite3.connect(target)
    try:
        source.backup(dest)
    finally:
        dest.close()
        source.close()
    return target


def pick_sample(conn):
    row = conn.execute(
        """
        SELECT country, COALESCE(NULLIF(category_l3_zh, ''), minor_category) AS category_key,
               MIN(year * 100 + month) AS start_int,
               MAX(year * 100 + month) AS end_int,
               COUNT(*) AS rows
        FROM sales_rows
        WHERE COALESCE(NULLIF(category_l3_zh, ''), minor_category) <> ''
        GROUP BY country, category_key
        ORDER BY rows DESC
        LIMIT 1
        """
    ).fetchone()
    if not row:
        return None
    return dict(row)


def int_to_month(value):
    year, month = divmod(int(value), 100)
    return f"{year:04d}-{month:02d}"


def benchmark(conn, sample, use_summary=False):
    start_month = int_to_month(sample["start_int"])
    end_month = int_to_month(sample["end_int"])
    start_int = month_int(start_month)
    end_int = month_int(end_month)
    if use_summary:
        monthly_sql = """
            SELECT country, year, month, category_key, SUM(sample_rows), SUM(sales), SUM(units)
            FROM monthly_category_sales
            WHERE country=? AND category_key=? AND (year * 100 + month) BETWEEN ? AND ?
            GROUP BY country, year, month, category_key
        """
    else:
        monthly_sql = """
            SELECT country, year, month,
                   COALESCE(NULLIF(category_l3_zh, ''), minor_category) AS category_key,
                   COUNT(*), SUM(sales), SUM(units)
            FROM sales_rows
            WHERE country=? AND COALESCE(NULLIF(category_l3_zh, ''), minor_category)=?
              AND (year * 100 + month) BETWEEN ? AND ?
            GROUP BY country, year, month, COALESCE(NULLIF(category_l3_zh, ''), minor_category)
        """
    top_seller_sql = """
        SELECT seller, SUM(sales) AS sales, SUM(units) AS units
        FROM sales_rows
        WHERE country=?
          AND (year * 100 + month) BETWEEN ? AND ?
          AND (minor_category=? OR category_l3_zh=?)
        GROUP BY seller
        ORDER BY sales DESC
        LIMIT 10
    """
    return {
        "monthly_aggregation": timed_query(conn, monthly_sql, (sample["country"], sample["category_key"], start_int, end_int)),
        "top_seller_query": timed_query(conn, top_seller_sql, (sample["country"], start_int, end_int, sample["category_key"], sample["category_key"])),
        "range": f"{start_month} 至 {end_month}",
    }


def main():
    backup = backup_database()
    report = {"backup": str(backup)}
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        sample = pick_sample(conn)
        report["sample"] = sample or {}
        if sample:
            report["before"] = benchmark(conn, sample, use_summary=False)
    init_db()
    started = time.perf_counter()
    ensure_monthly_summary()
    report["migration_ms"] = round((time.perf_counter() - started) * 1000, 2)
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        report["summary_rows"] = conn.execute("SELECT COUNT(*) AS n FROM monthly_category_sales").fetchone()["n"]
        if report.get("sample"):
            report["after"] = benchmark(conn, report["sample"], use_summary=True)
    report_path = DATA_DIR / "备份" / f"月度汇总迁移报告-{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({**report, "report": str(report_path)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
