import tempfile
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from convert import CONVERTERS

MAX_BYTES = 50 * 1024 * 1024
MIME = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
    "csv": "text/csv; charset=utf-8",
    "json": "application/json",
}


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        query = parse_qs(urlparse(self.path).query)
        name = Path(query.get("name", [""])[0]).name
        ext = Path(name).suffix.lower()
        if ext not in CONVERTERS:
            return self.error(400, "Upload a .pdf, .xlsx or .xlsm file")

        fmt = query.get("format", [next(iter(CONVERTERS[ext]))])[0]
        if fmt not in CONVERTERS[ext]:
            return self.error(400, f"Cannot make .{fmt} from {ext}")

        size = int(self.headers.get("Content-Length") or 0)
        if not 0 < size <= MAX_BYTES:
            return self.error(413, "File is empty or larger than 50 MB")

        with tempfile.TemporaryDirectory() as directory:
            src = Path(directory) / f"in{ext}"
            dst = Path(directory) / f"out.{fmt}"
            src.write_bytes(self.rfile.read(size))
            try:
                CONVERTERS[ext][fmt](str(src), str(dst))
                data = dst.read_bytes()
            except Exception as exc:
                return self.error(422, f"Could not convert: {exc}")

        self.send_response(200)
        self.send_header("Content-Type", MIME[fmt])
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{Path(name).stem}.{fmt}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        return self.error(405, "Use POST /convert")

    def error(self, code, message):
        body = message.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
