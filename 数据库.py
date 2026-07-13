import sqlite3
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "用户数据"
DB_PATH = DATA_DIR / "gmv诊断系统.sqlite3"


def get_conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.executescript(
            """
            PRAGMA journal_mode=WAL;

            CREATE TABLE IF NOT EXISTS raw_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL UNIQUE,
                country TEXT NOT NULL,
                year INTEGER,
                month INTEGER,
                row_count INTEGER DEFAULT 0,
                imported_rows INTEGER DEFAULT 0,
                notes TEXT DEFAULT '',
                imported_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS sales_rows (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER,
                country TEXT NOT NULL,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                asin TEXT,
                sku TEXT,
                title TEXT,
                brand TEXT,
                category_path TEXT,
                major_category TEXT,
                minor_category TEXT,
                sales REAL DEFAULT 0,
                units REAL DEFAULT 0,
                price REAL DEFAULT 0,
                seller TEXT,
                seller_location TEXT,
                seller_info TEXT,
                seller_home TEXT,
                product_url TEXT,
                image_url TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_sales_country_month
            ON sales_rows(country, year, month);

            CREATE INDEX IF NOT EXISTS idx_sales_category
            ON sales_rows(country, year, minor_category);

            CREATE INDEX IF NOT EXISTS idx_sales_country_category_month
            ON sales_rows(country, category_l3_zh, year, month);

            CREATE INDEX IF NOT EXISTS idx_sales_country_minor_month
            ON sales_rows(country, minor_category, year, month);

            CREATE INDEX IF NOT EXISTS idx_sales_country_month_seller
            ON sales_rows(country, year, month, seller);

            CREATE TABLE IF NOT EXISTS monthly_category_sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT NOT NULL,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                category_key TEXT NOT NULL,
                category_path TEXT DEFAULT '',
                category_l1_zh TEXT DEFAULT '',
                category_l2_zh TEXT DEFAULT '',
                category_l3_zh TEXT DEFAULT '',
                sample_rows INTEGER DEFAULT 0,
                sales REAL DEFAULT 0,
                units REAL DEFAULT 0,
                updated_at TEXT NOT NULL,
                UNIQUE(country, year, month, category_key)
            );

            CREATE INDEX IF NOT EXISTS idx_monthly_country_category_month
            ON monthly_category_sales(country, category_key, year, month);

            CREATE INDEX IF NOT EXISTS idx_monthly_country_month
            ON monthly_category_sales(country, year, month);

            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zh_cate2 TEXT,
                zh_cate3 TEXT,
                zh_cate4 TEXT,
                en_cate2 TEXT,
                en_cate3 TEXT,
                en_cate4 TEXT,
                source TEXT DEFAULT '庭院类目tree'
            );

            CREATE TABLE IF NOT EXISTS category_tree (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                level1_en TEXT DEFAULT '',
                level1_es TEXT DEFAULT '',
                level1_de TEXT DEFAULT '',
                level1_zh TEXT DEFAULT '',
                level2_en TEXT DEFAULT '',
                level2_es TEXT DEFAULT '',
                level2_de TEXT DEFAULT '',
                level2_zh TEXT DEFAULT '',
                level3_en TEXT DEFAULT '',
                level3_es TEXT DEFAULT '',
                level3_de TEXT DEFAULT '',
                level3_zh TEXT DEFAULT '',
                source TEXT DEFAULT '多语言类目树'
            );

            CREATE TABLE IF NOT EXISTS holidays (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rule_id INTEGER,
                country TEXT NOT NULL,
                year INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                name_cn TEXT NOT NULL,
                name_local TEXT DEFAULT '',
                kind TEXT NOT NULL,
                consumer_note TEXT DEFAULT '',
                prep_days INTEGER DEFAULT 120,
                color TEXT DEFAULT 'blue',
                generated_from TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS holiday_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT NOT NULL,
                name_cn TEXT NOT NULL,
                name_local TEXT DEFAULT '',
                rule_text TEXT NOT NULL,
                is_floating INTEGER DEFAULT 0,
                consumer_note TEXT DEFAULT '',
                kind TEXT DEFAULT '节日日历',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(country, name_cn)
            );

            CREATE TABLE IF NOT EXISTS analysis_conclusions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT NOT NULL,
                year INTEGER DEFAULT 0,
                time_range_start TEXT DEFAULT '',
                time_range_end TEXT DEFAULT '',
                category_key TEXT NOT NULL,
                category_path TEXT DEFAULT '',
                sample_months TEXT NOT NULL,
                avg_sales REAL DEFAULT 0,
                peak_months TEXT NOT NULL,
                price_low REAL DEFAULT 0,
                price_high REAL DEFAULT 0,
                top_sellers TEXT NOT NULL,
                chart_data TEXT NOT NULL,
                completeness_note TEXT DEFAULT '',
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS analysis_cache (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT NOT NULL,
                category TEXT NOT NULL,
                analysis_type TEXT NOT NULL CHECK(analysis_type IN ('year', 'all')),
                year INTEGER,
                peak_months TEXT NOT NULL,
                price_range_low REAL DEFAULT 0,
                price_range_high REAL DEFAULT 0,
                top10_sellers TEXT NOT NULL,
                data_range_start TEXT NOT NULL,
                data_range_end TEXT NOT NULL,
                calc_time TEXT NOT NULL,
                is_valid INTEGER DEFAULT 1,
                category_path TEXT DEFAULT '',
                category_l1_zh TEXT DEFAULT '',
                category_l2_zh TEXT DEFAULT '',
                category_l3_zh TEXT DEFAULT '',
                sample_months TEXT DEFAULT '[]',
                avg_sales REAL DEFAULT 0,
                chart_data TEXT DEFAULT '[]',
                completeness_note TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS merchant_resources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT NOT NULL,
                category_key TEXT NOT NULL,
                seller_name TEXT NOT NULL,
                company_name TEXT DEFAULT '',
                contact TEXT DEFAULT '',
                address TEXT DEFAULT '',
                email TEXT DEFAULT '',
                seller_location TEXT DEFAULT '',
                seller_info TEXT DEFAULT '',
                seller_home TEXT DEFAULT '',
                sales REAL DEFAULT 0,
                units REAL DEFAULT 0,
                source TEXT DEFAULT '销售数据Top10',
                notes TEXT DEFAULT '',
                updated_at TEXT NOT NULL,
                UNIQUE(country, category_key, seller_name)
            );

            CREATE INDEX IF NOT EXISTS idx_analysis_country_category_range
            ON analysis_conclusions(country, category_key, time_range_start, time_range_end);

            CREATE INDEX IF NOT EXISTS idx_analysis_cache_lookup
            ON analysis_cache(country, category, analysis_type, year, is_valid);

            CREATE INDEX IF NOT EXISTS idx_analysis_cache_valid_type
            ON analysis_cache(is_valid, analysis_type, year);

            CREATE INDEX IF NOT EXISTS idx_merchants_country_category_sales
            ON merchant_resources(country, category_key, sales);

            CREATE TABLE IF NOT EXISTS wisdom (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                country TEXT DEFAULT '',
                category_key TEXT DEFAULT '',
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                keywords TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS todos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                text TEXT NOT NULL,
                done INTEGER DEFAULT 0,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS chat_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS translations (
                original_name TEXT PRIMARY KEY,
                zh_name TEXT DEFAULT '',
                updated_at TEXT NOT NULL
            );
            """
        )
        ensure_column(conn, "merchant_resources", "seller_info", "TEXT DEFAULT ''")
        ensure_column(conn, "holidays", "rule_id", "INTEGER")
        ensure_column(conn, "holidays", "generated_from", "TEXT DEFAULT ''")
        ensure_column(conn, "wisdom", "country", "TEXT DEFAULT ''")
        ensure_column(conn, "wisdom", "updated_at", "TEXT DEFAULT ''")
        migrate_analysis_conclusions(conn)
        ensure_column(conn, "analysis_conclusions", "source", "TEXT DEFAULT '销售数据'")
        ensure_column(conn, "analysis_conclusions", "source_ref", "TEXT DEFAULT ''")
        ensure_column(conn, "analysis_conclusions", "time_range_start", "TEXT DEFAULT ''")
        ensure_column(conn, "analysis_conclusions", "time_range_end", "TEXT DEFAULT ''")
        ensure_column(conn, "sales_rows", "category_l1_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "sales_rows", "category_l2_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "sales_rows", "category_l3_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "sales_rows", "category_match_method", "TEXT DEFAULT ''")
        ensure_column(conn, "analysis_conclusions", "category_l1_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "analysis_conclusions", "category_l2_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "analysis_conclusions", "category_l3_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "merchant_resources", "category_l1_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "merchant_resources", "category_l2_zh", "TEXT DEFAULT ''")
        ensure_column(conn, "merchant_resources", "category_l3_zh", "TEXT DEFAULT ''")
        conn.execute("UPDATE wisdom SET updated_at=created_at WHERE COALESCE(updated_at, '')=''")
        conn.execute("UPDATE analysis_conclusions SET source='销售数据' WHERE COALESCE(source, '')=''")


def ensure_column(conn, table, column, definition):
    existing = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}
    if column not in existing:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def migrate_analysis_conclusions(conn):
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='analysis_conclusions'"
    ).fetchone()
    table_sql = row["sql"] if row else ""
    columns = {r["name"] for r in conn.execute("PRAGMA table_info(analysis_conclusions)")}
    needs_rebuild = (
        "time_range_start" not in columns
        or "time_range_end" not in columns
        or "UNIQUE(country, year, category_key)" in table_sql
    )
    if not needs_rebuild:
        return

    conn.execute("DROP TABLE IF EXISTS analysis_conclusions_new")
    conn.execute(
        """
        CREATE TABLE analysis_conclusions_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            country TEXT NOT NULL,
            year INTEGER DEFAULT 0,
            time_range_start TEXT DEFAULT '',
            time_range_end TEXT DEFAULT '',
            category_key TEXT NOT NULL,
            category_path TEXT DEFAULT '',
            sample_months TEXT NOT NULL,
            avg_sales REAL DEFAULT 0,
            peak_months TEXT NOT NULL,
            price_low REAL DEFAULT 0,
            price_high REAL DEFAULT 0,
            top_sellers TEXT NOT NULL,
            chart_data TEXT NOT NULL,
            completeness_note TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            source TEXT DEFAULT '销售数据',
            source_ref TEXT DEFAULT '',
            category_l1_zh TEXT DEFAULT '',
            category_l2_zh TEXT DEFAULT '',
            category_l3_zh TEXT DEFAULT ''
        )
        """
    )

    def expr(column, fallback):
        return column if column in columns else fallback

    time_start = "time_range_start" if "time_range_start" in columns else "''"
    time_end = "time_range_end" if "time_range_end" in columns else "''"
    conn.execute(
        f"""
        INSERT INTO analysis_conclusions_new(
            id, country, year, time_range_start, time_range_end, category_key, category_path,
            sample_months, avg_sales, peak_months, price_low, price_high, top_sellers,
            chart_data, completeness_note, created_at, source, source_ref,
            category_l1_zh, category_l2_zh, category_l3_zh
        )
        SELECT
            {expr('id', 'NULL')},
            {expr('country', "''")},
            {expr('year', '0')},
            {time_start},
            {time_end},
            {expr('category_key', "''")},
            {expr('category_path', "''")},
            {expr('sample_months', "'[]'")},
            {expr('avg_sales', '0')},
            {expr('peak_months', "'[]'")},
            {expr('price_low', '0')},
            {expr('price_high', '0')},
            {expr('top_sellers', "'[]'")},
            {expr('chart_data', "'[]'")},
            {expr('completeness_note', "''")},
            {expr('created_at', "datetime('now')")},
            {expr('source', "'销售数据'")},
            {expr('source_ref', "''")},
            {expr('category_l1_zh', "''")},
            {expr('category_l2_zh', "''")},
            {expr('category_l3_zh', "''")}
        FROM analysis_conclusions
        """
    )
    conn.execute("DROP TABLE analysis_conclusions")
    conn.execute("ALTER TABLE analysis_conclusions_new RENAME TO analysis_conclusions")
