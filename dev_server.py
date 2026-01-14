#!/usr/bin/env python3
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
import urllib.error
import urllib.request
from urllib.parse import parse_qs, urlparse

MONDAY_API_URL = "https://api.monday.com/v2"
PORT = int(os.environ.get("PORT", "8000"))


class ProxyHandler(SimpleHTTPRequestHandler):
    def send_cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header(
            "Access-Control-Allow-Headers", "Content-Type, Authorization"
        )
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")

    def do_OPTIONS(self):
        if self.path != "/api/monday":
            self.send_error(404, "Not found")
            return
        self.send_response(204)
        self.send_cors_headers()
        self.end_headers()

    def do_POST(self):
        if self.path != "/api/monday":
            self.send_error(404, "Not found")
            return
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length) if length > 0 else b""
        token = self.headers.get("Authorization", "")
        if not token:
            self.send_response(400)
            self.send_cors_headers()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"error":"Missing Authorization header."}')
            return

        request = urllib.request.Request(
            MONDAY_API_URL,
            data=body,
            method="POST",
            headers={
                "Content-Type": "application/json",
                "Authorization": token,
            },
        )

        try:
            with urllib.request.urlopen(request) as response:
                status = response.status
                payload = response.read()
                content_type = response.headers.get(
                    "Content-Type", "application/json"
                )
        except urllib.error.HTTPError as exc:
            status = exc.code
            payload = exc.read()
            content_type = exc.headers.get("Content-Type", "application/json")
        except Exception as exc:
            self.send_response(502)
            self.send_cors_headers()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            message = json.dumps({"error": f"Proxy error: {exc}"}).encode("utf-8")
            self.wfile.write(message)
            return

        self.send_response(status)
        self.send_cors_headers()
        self.send_header("Content-Type", content_type)
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if self.path.startswith("/api/pto"):
            parsed = urlparse(self.path)
            params = parse_qs(parsed.query)
            target_url = params.get("url", [None])[0]
            if not target_url:
                self.send_response(400)
                self.send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"error":"Missing url parameter."}')
                return
            if not target_url.startswith(("http://", "https://")):
                self.send_response(400)
                self.send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"error":"Invalid url protocol."}')
                return

            request = urllib.request.Request(target_url, method="GET")
            try:
                with urllib.request.urlopen(request) as response:
                    status = response.status
                    payload = response.read()
                    content_type = response.headers.get(
                        "Content-Type", "application/json"
                    )
            except urllib.error.HTTPError as exc:
                status = exc.code
                payload = exc.read()
                content_type = exc.headers.get("Content-Type", "application/json")
            except Exception as exc:
                self.send_response(502)
                self.send_cors_headers()
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                message = json.dumps({"error": f"Proxy error: {exc}"}).encode("utf-8")
                self.wfile.write(message)
                return

            self.send_response(status)
            self.send_cors_headers()
            self.send_header("Content-Type", content_type)
            self.end_headers()
            self.wfile.write(payload)
            return

        super().do_GET()


def main():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), ProxyHandler)
    print(f"Serving on http://localhost:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
