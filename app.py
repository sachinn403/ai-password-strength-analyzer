"""Local web server for the AI-powered password strength analyzer."""

from __future__ import annotations

import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from password_ai import PasswordStrengthAnalyzer, generate_secure_password


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
ANALYZER = PasswordStrengthAnalyzer.from_default_model()


class PasswordAnalyzerHandler(SimpleHTTPRequestHandler):
    """Serves the frontend and handles password analysis requests."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self) -> None:
        route = urlparse(self.path).path
        if route not in {"/api/analyze", "/api/generate"}:
            self._send_json({"error": "Not found"}, status=404)
            return

        try:
            payload = self._read_json_payload()
            if route == "/api/generate":
                result = self._generate_password(payload)
            else:
                password = str(payload.get("password", ""))
                context = payload.get("context") or {}
                result = ANALYZER.analyze(password, context=context)
            self._send_json(result)
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON request."}, status=400)
        except Exception as exc:
            self._send_json({"error": f"Analysis failed: {exc}"}, status=500)

    def _read_json_payload(self) -> dict:
        size = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(size).decode("utf-8") or "{}")

    def _generate_password(self, payload: dict) -> dict:
        length = int(payload.get("length", 18))
        password = generate_secure_password(length)
        return {
            "password": password,
            "length": len(password),
            "analysis": ANALYZER.analyze(password),
        }

    def _send_json(self, payload: dict, status: int = 200) -> None:
        encoded = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)


def run(host: str = "0.0.0.0", port: int = 8000) -> None:
    server = ThreadingHTTPServer((host, port), PasswordAnalyzerHandler)
    print(f"AI Password Strength Analyzer running at http://{host}:{port}")
    print("Press Ctrl+C to stop the server.")
    server.serve_forever()


if __name__ == "__main__":
    run()
