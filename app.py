"""Single-user loopback demo, deliberately serial to avoid workbook write races."""
import argparse
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import secrets
from urllib.parse import urlparse, parse_qs

from core import BASE, DemoStore

def create_server(port=8765, storage=None):
    store = DemoStore(storage or BASE / "runtime")
    token = secrets.token_urlsafe(32)

    class Handler(BaseHTTPRequestHandler):
        def send(self, data, content_type="application/json; charset=utf-8", status=200, download=False):
            if not isinstance(data, bytes):
                data = json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Content-Security-Policy", "default-src 'self'; style-src 'self'; script-src 'self'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'")
            if download:
                self.send_header("Content-Disposition", 'attachment; filename="demo.xlsx"')
            self.end_headers()
            self.wfile.write(data)

        def allowed_host(self):
            return self.headers.get("Host") == f"127.0.0.1:{self.server.server_port}"

        def do_GET(self):
            if not self.allowed_host():
                return self.send({"error":"127.0.0.1 주소로 접속하세요."}, status=403)
            url = urlparse(self.path)
            try:
                static = {"/":("static/index.html","text/html; charset=utf-8"),
                          "/app.js":("static/app.js","text/javascript; charset=utf-8"),
                          "/style.css":("static/style.css","text/css; charset=utf-8"),
                          "/plan.md":("docs/준비계획.md","text/plain; charset=utf-8"),
                          "/readme.md":("README.md","text/plain; charset=utf-8"),
                          "/script.md":("docs/시연대본.md","text/plain; charset=utf-8")}
                if url.path in static:
                    file, mime = static[url.path]
                    return self.send((BASE/file).read_bytes(),mime)
                if url.path == "/api/config":
                    return self.send({"token":token, "mode":"offline_fixture", "app":"macro-career-lab-v1"})
                if url.path == "/api/plan":
                    return self.send((BASE/"data/plan.json").read_bytes())
                if url.path == "/api/history":
                    return self.send(store.history())
                if url.path == "/api/evidence":
                    path = BASE/"evidence/results.json"
                    return self.send(json.loads(path.read_text(encoding="utf-8")) if path.exists() else {"cases":[],"notice":"검증 결과가 아직 없습니다."})
                if url.path in {"/api/state","/download"}:
                    run_id = parse_qs(url.query).get("id",[""])[0]
                    if url.path == "/api/state":
                        return self.send(store.state(run_id))
                    return self.send((store.directory(run_id)/"demo.xlsx").read_bytes(),
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",download=True)
                return self.send({"error":"해당 페이지가 없습니다."},status=404)
            except (ValueError, FileNotFoundError, KeyError) as exc:
                self.send({"error":str(exc)},status=400)

        def do_POST(self):
            origin = self.headers.get("Origin")
            if (not self.allowed_host() or self.headers.get("X-Demo-Token") != token or
                (origin is not None and origin != f"http://127.0.0.1:{self.server.server_port}")):
                return self.send({"error":"로컬 데모 화면에서 다시 실행하세요."},status=403)
            try:
                length = int(self.headers.get("Content-Length","0"))
                if not 0 < length <= 4096:
                    return self.send({"error":"요청 크기가 올바르지 않습니다."},status=400)
                body = json.loads(self.rfile.read(length))
                if not isinstance(body,dict):
                    raise ValueError("JSON 객체가 필요합니다.")
                route = urlparse(self.path).path
                if route == "/api/run":
                    result = store.run(body.get("scenario","baseline"))
                elif route == "/api/review":
                    result = store.review(body.get("id"),body.get("score"))
                elif route == "/api/reapply":
                    result = store.reapply(body.get("id"))
                else:
                    return self.send({"error":"지원하지 않는 동작입니다."},status=404)
                self.send(result)
            except (ValueError,KeyError,TypeError) as exc:
                self.send({"error":str(exc)},status=400)
            except Exception:
                self.send({"error":"실행 실패. 로컬 실행 로그와 파일이 열려 있는지 확인하세요."},status=500)

    return HTTPServer(("127.0.0.1",port),Handler)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port",type=int,default=8765)
    args = parser.parse_args()
    server = create_server(args.port)
    print(f"Demo: http://127.0.0.1:{server.server_port}",flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
