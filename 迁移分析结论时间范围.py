#!/usr/bin/env python3
import json
import sqlite3
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DEFAULT_DB = ROOT / "用户数据" / "gmv诊断系统.sqlite3"


def table_columns(conn, table):
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}


def ensure_column(conn, table, column, definition):
    columns = table_columns(conn, table)
    if column in columns:
        return False
    conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
    return True


def backup_database(db_path):
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_dir = db_path.parent / "备份"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir / f"gmv诊断系统-时间范围迁移前备份-{timestamp}.sqlite3"
    source = sqlite3.connect(db_path)
    try:
        target = sqlite3.connect(backup_path)
        try:
            source.backup(target)
        finally:
            target.close()
    finally:
        source.close()
    return backup_path


def year_to_range(year_value):
    try:
        year = int(year_value)
    except (TypeError, ValueError):
        return None, None
    if year <= 0:
        return None, None
    return f"{year:04d}-01", f"{year:04d}-12"


def migrate(db_path):
    db_path = Path(db_path).expanduser().resolve()
    if not db_path.exists():
        raise FileNotFoundError(f"数据库不存在：{db_path}")

    backup_path = backup_database(db_path)
    report = {
        "database": str(db_path),
        "backup": str(backup_path),
        "columns_added": [],
        "total_records": 0,
        "already_had_time_range": 0,
        "records_needing_migration": 0,
        "migrated_records": 0,
        "failed_records": [],
    }

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("BEGIN")
        if ensure_column(conn, "analysis_conclusions", "time_range_start", "TEXT DEFAULT ''"):
            report["columns_added"].append("time_range_start")
        if ensure_column(conn, "analysis_conclusions", "time_range_end", "TEXT DEFAULT ''"):
            report["columns_added"].append("time_range_end")

        rows = conn.execute(
            """
            SELECT id, year, time_range_start, time_range_end
            FROM analysis_conclusions
            ORDER BY id
            """
        ).fetchall()
        report["total_records"] = len(rows)

        for row in rows:
            has_start = bool(row["time_range_start"])
            has_end = bool(row["time_range_end"])
            if has_start and has_end:
                report["already_had_time_range"] += 1
                continue

            report["records_needing_migration"] += 1
            start, end = year_to_range(row["year"])
            if not start or not end:
                report["failed_records"].append(
                    {"id": row["id"], "year": row["year"], "reason": "year 为空或不是有效年份，无法推导时间范围"}
                )
                continue

            conn.execute(
                """
                UPDATE analysis_conclusions
                SET time_range_start = ?, time_range_end = ?
                WHERE id = ?
                """,
                (row["time_range_start"] or start, row["time_range_end"] or end, row["id"]),
            )
            report["migrated_records"] += 1

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    report_path = backup_path.with_name(backup_path.stem.replace("备份", "报告") + ".json")
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    report["report"] = str(report_path)
    return report


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_DB
    result = migrate(target)
    print(json.dumps(result, ensure_ascii=False, indent=2))
