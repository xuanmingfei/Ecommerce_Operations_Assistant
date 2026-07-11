import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

import cgi
import csv
import json
import mimetypes
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from 分析引擎 import (
    DATA_DIR,
    ROOT,
    add_wisdom,
    add_holiday_rule,
    build_analysis_notifications,
    add_merchant,
    delete_holiday_rule,
    delete_merchant,
    delete_wisdom,
    export_translation_terms,
    get_state,
    import_merchants_file,
    import_sales_file,
    import_translation_file,
    local_chat_answer,
    month_key,
    rebuild_analysis,
    seed_initial_data,
    update_holiday_rule,
    update_merchant,
    update_wisdom,
)
from 数据库 import init_db


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path):
        path = Path(path)
        if not path.exists() or not path.is_file():
            self.send_error(404)
            return
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_download(self, path, filename=None):
        path = Path(path)
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename or path.name)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_csv_download(self, rows, filename):
        target = tempfile.NamedTemporaryFile(delete=False, suffix=".csv", dir=DATA_DIR)
        target.close()
        path = Path(target.name)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(rows)
        self.send_download(path, filename)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        return json.loads(raw or "{}")

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        if path == "/":
            self.send_file(ROOT / "页面" / "首页.html")
            return
        if path == "/static/样式.css":
            self.send_file(ROOT / "页面" / "样式.css")
            return
        if path == "/static/交互.js":
            self.send_file(ROOT / "页面" / "交互.js")
            return
        if path == "/api/state":
            qs = parse_qs(parsed.query)
            self.send_json(get_state(qs.get("year", [None])[0], qs.get("month", [None])[0]))
            return
        if path == "/api/merchant/template":
            self.send_csv_download([["店铺名字", "公司名称", "联系电话", "邮箱", "经营地址", "主营三级类目", "备注"]], "商家批量导入模板.csv")
            return
        if path == "/api/translations/export":
            self.send_download(export_translation_terms(), "待翻译词汇.xlsx")
            return
        self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/chat":
                payload = self.read_json()
                self.send_json(local_chat_answer(payload.get("message", "")))
                return
            if parsed.path == "/api/wisdom/add":
                payload = self.read_json()
                add_wisdom(payload.get("title", ""), payload.get("content", ""), payload.get("category_key", ""), payload.get("keywords", ""), payload.get("country", ""))
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/wisdom/update":
                payload = self.read_json()
                update_wisdom(payload.get("id"), payload)
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/wisdom/delete":
                payload = self.read_json()
                delete_wisdom(payload.get("id"))
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/holiday/add":
                payload = self.read_json()
                add_holiday_rule(payload)
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/holiday/update":
                payload = self.read_json()
                update_holiday_rule(payload.get("id"), payload)
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/holiday/delete":
                payload = self.read_json()
                delete_holiday_rule(payload.get("id"))
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/merchant/add":
                payload = self.read_json()
                add_merchant(payload)
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/merchant/update":
                payload = self.read_json()
                update_merchant(payload.get("id"), payload)
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/merchant/delete":
                payload = self.read_json()
                delete_merchant(payload.get("id"))
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/merchant/import":
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
                file_item = form["file"] if "file" in form else None
                if file_item is None or not file_item.filename:
                    self.send_json({"ok": False, "message": "没有收到文件"}, 400)
                    return
                suffix = Path(file_item.filename).suffix or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=DATA_DIR) as tmp:
                    tmp.write(file_item.file.read())
                    tmp_path = Path(tmp.name)
                self.send_json(import_merchants_file(tmp_path, form.getfirst("country", "")))
                return
            if parsed.path == "/api/translations/import":
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
                file_item = form["file"] if "file" in form else None
                if file_item is None or not file_item.filename:
                    self.send_json({"ok": False, "message": "没有收到文件"}, 400)
                    return
                suffix = Path(file_item.filename).suffix or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=DATA_DIR) as tmp:
                    tmp.write(file_item.file.read())
                    tmp_path = Path(tmp.name)
                self.send_json(import_translation_file(tmp_path))
                return
            if parsed.path == "/api/rebuild":
                rebuild_analysis()
                self.send_json({"ok": True})
                return
            if parsed.path == "/api/upload":
                form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
                file_item = form["file"] if "file" in form else None
                country = form.getfirst("country", "")
                start_month = form.getfirst("time_range_start", "")
                end_month = form.getfirst("time_range_end", "")
                if file_item is None or not file_item.filename:
                    self.send_json({"ok": False, "message": "没有收到文件"}, 400)
                    return
                if not country.strip():
                    self.send_json({"ok": False, "message": "请先选择或输入数据归属国家"}, 400)
                    return
                suffix = Path(file_item.filename).suffix or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=DATA_DIR) as tmp:
                    tmp.write(file_item.file.read())
                    tmp_path = Path(tmp.name)
                target = DATA_DIR / "临时上传" / file_item.filename
                target.parent.mkdir(parents=True, exist_ok=True)
                tmp_path.replace(target)
                result = import_sales_file(target, country=country, source_copy=True)
                if result.get("ok"):
                    uploaded_month = month_key(result["year"], result["month"]) if result.get("year") and result.get("month") else ""
                    analysis_start = start_month or uploaded_month
                    analysis_end = end_month or uploaded_month
                    rebuild_result = rebuild_analysis(analysis_start, analysis_end, country=result.get("country"))
                    result.update({k: v for k, v in rebuild_result.items() if k in ("time_range_start", "time_range_end")})
                    result["notifications"] = build_analysis_notifications(
                        result.get("country"),
                        result.get("time_range_start") or analysis_start,
                        result.get("time_range_end") or analysis_end,
                        result.get("categories", []),
                    )
                self.send_json(result)
                return
        except Exception as exc:
            self.send_json({"ok": False, "message": str(exc)}, 500)
            return
        self.send_error(404)


def run():
    init_db()
    seed_initial_data()
    port = 8780
    while True:
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except PermissionError as exc:
            raise RuntimeError("当前环境不允许直接启动本地服务，请用 Codex 授权方式启动。") from exc
        except OSError:
            port += 1
            if port > 65535:
                raise RuntimeError("没有找到可用端口，请先关闭旧的 GMV 诊断系统服务。")
    print(f"GMV诊断系统已启动：http://127.0.0.1:{port}/")
    print("按 Control+C 可以关闭。")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nGMV诊断系统已关闭。")


if __name__ == "__main__":
    run()
