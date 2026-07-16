from __future__ import annotations
import json, os, pathlib, urllib.error, urllib.request
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

ROOT = pathlib.Path(__file__).resolve().parent
CONFIG_PATHS = [ROOT / 'gemini.config.json', ROOT / 'gemini.config.local.json']
RETRYABLE_STATUS = {404, 429, 500, 502, 503, 504}


def load_config() -> dict:
    for path in CONFIG_PATHS:
        if path.exists():
            try:
                return json.loads(path.read_text(encoding='utf-8'))
            except Exception:
                return {}
    return {}


def normalize_model_list(*groups):
    result = []
    seen = set()
    for group in groups:
        if not group:
            continue
        if isinstance(group, str):
            group = [group]
        for item in group:
            model = str(item or '').strip()
            if not model or model in seen:
                continue
            if any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._-' for c in model):
                continue
            seen.add(model)
            result.append(model)
    return result


def call_gemini(model: str, key: str, request_body: dict):
    data = json.dumps(request_body, separators=(',', ':'), ensure_ascii=False).encode('utf-8')
    req = urllib.request.Request(
        f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent',
        data=data,
        headers={'Content-Type': 'application/json', 'x-goog-api-key': key},
        method='POST',
    )
    try:
        with urllib.request.urlopen(req, timeout=240) as res:
            return res.status, res.read(), None
    except urllib.error.HTTPError as err:
        return err.code, err.read(), err


class Handler(SimpleHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        if self.path == '/api/gemini-config':
            cfg = load_config()
            body = json.dumps({
                'hasApiKey': bool(str(cfg.get('apiKey') or os.getenv('GEMINI_API_KEY') or '').strip()),
                'primaryModel': str(cfg.get('primaryModel') or 'gemini-2.5-flash'),
                'fallbackModels': [str(x) for x in cfg.get('fallbackModels') or [] if str(x).strip()],
                'allowBrowserOverride': bool(cfg.get('allowBrowserOverride', True)),
            }).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()

    def do_POST(self):
        if self.path != '/api/gemini':
            self.send_error(404)
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if length <= 0 or length > 80 * 1024 * 1024:
                raise ValueError('Ukuran permintaan tidak valid.')
            payload = json.loads(self.rfile.read(length).decode('utf-8'))
            cfg = load_config()
            key = str(payload.get('apiKey') or cfg.get('apiKey') or os.getenv('GEMINI_API_KEY') or '').strip()
            request_body = payload.get('request')
            if not key:
                raise ValueError('Gemini API key belum diisi pada gemini.config.json, gemini.config.local.json, atau browser.')
            if not isinstance(request_body, dict):
                raise ValueError('Body permintaan Gemini tidak valid.')

            preferred = payload.get('modelChain')
            if not preferred:
                preferred = normalize_model_list(payload.get('model'), payload.get('fallbackModels'))
            models = normalize_model_list(preferred, cfg.get('primaryModel'), cfg.get('fallbackModels'))
            if not models:
                raise ValueError('Model Gemini belum diatur.')

            attempts = []
            final_status, final_body = 500, b''
            for idx, model in enumerate(models):
                status, body, err = call_gemini(model, key, request_body)
                attempts.append({'model': model, 'status': status})
                if 200 <= status < 300:
                    final_status, final_body = status, body
                    self.send_response(final_status)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Content-Length', str(len(final_body)))
                    self.send_header('X-Gemini-Model-Used', model)
                    self.send_header('X-Gemini-Attempts', json.dumps(attempts, ensure_ascii=False))
                    self.end_headers()
                    self.wfile.write(final_body)
                    return
                # Jangan retry untuk auth error atau request error yang jelas bukan kuota/kapasitas
                if status in (400, 401, 403):
                    final_status, final_body = status, body
                    break
                if status not in RETRYABLE_STATUS and idx == len(models) - 1:
                    final_status, final_body = status, body
                    break
                if status not in RETRYABLE_STATUS and idx < len(models) - 1:
                    # Tetap coba model berikutnya bila model bermasalah/unsupported.
                    final_status, final_body = status, body
                    continue
                final_status, final_body = status, body

            self.send_response(final_status)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(final_body)))
            self.send_header('X-Gemini-Attempts', json.dumps(attempts, ensure_ascii=False))
            self.end_headers()
            self.wfile.write(final_body)
        except Exception as exc:
            body = json.dumps({'error': {'message': str(exc)}}).encode('utf-8')
            self.send_response(400)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)


if __name__ == '__main__':
    for port in range(8765, 8786):
        try:
            server = ThreadingHTTPServer((os.getenv('HOST','0.0.0.0'), int(os.getenv('PORT', port))), Handler)
            break
        except OSError:
            continue
    else:
        raise SystemExit('Tidak ada port kosong.')
    import webbrowser
    url = f'http://127.0.0.1:{port}/index.html'
    print('Generator MCU:', url)
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
