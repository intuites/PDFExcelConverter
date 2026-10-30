"""Web UI for convert.py. Run: python server.py  then open http://127.0.0.1:8000"""
import tempfile
import socket
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
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


class LocalServer(ThreadingHTTPServer):
    # Windows permits multiple SO_REUSEADDR listeners on one port. Requests
    # can then reach an old parser despite starting a newer server instance.
    allow_reuse_address = False

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


class Handler(SimpleHTTPRequestHandler):
    def do_POST(self):
        url = urlparse(self.path)
        if url.path != "/convert":
            return self.send_error(404)
        query = parse_qs(url.query)
        # Path(...).name strips any directory parts a client might send
        name = Path(query.get("name", [""])[0]).name
        ext = Path(name).suffix.lower()
        size = int(self.headers.get("Content-Length") or 0)
        if ext not in CONVERTERS:
            return self.send_error(400, "Upload a .pdf, .xls, .xlsx or .xlsm file")
        fmt = query.get("format", [next(iter(CONVERTERS[ext]))])[0]
        if fmt not in CONVERTERS[ext]:
            return self.send_error(400, f"Cannot make .{fmt} from {ext}")
        if not 0 < size <= MAX_BYTES:
            return self.send_error(413, "File is empty or larger than 50 MB")

        with tempfile.TemporaryDirectory() as d:
            src, dst = Path(d) / f"in{ext}", Path(d) / f"out.{fmt}"
            src.write_bytes(self.rfile.read(size))
            try:
                CONVERTERS[ext][fmt](str(src), str(dst))
            except Exception as e:  # corrupt or unsupported file: report it instead of dropping the connection
                return self.send_error(422, f"Could not convert: {e}")
            data = dst.read_bytes()

        self.send_response(200)
        self.send_header("Content-Type", MIME[fmt])
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{Path(name).stem}.{fmt}"')
        self.end_headers()
        self.wfile.write(data)


if __name__ == "__main__":
    web = Path(__file__).parent / "web"
    # 127.0.0.1 only: the server is not reachable from other machines
    server = LocalServer(("127.0.0.1", 8000), partial(Handler, directory=str(web)))
    print("Open http://127.0.0.1:8000")
    server.serve_forever()
