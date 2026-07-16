from __future__ import annotations
import base64, io, json, os, pathlib, re, urllib.error, urllib.request
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Mm, Pt

ROOT = pathlib.Path(__file__).resolve().parent
RETRYABLE_STATUS = {404, 429, 500, 502, 503, 504}
DEFAULT_PRIMARY_MODEL = 'gemini-2.5-flash'
DEFAULT_FALLBACK_MODELS = ['gemini-2.5-flash-lite', 'gemini-2.0-flash']


def normalize_model_list(*groups):
    result = []
    seen = set()
    for group in groups:
        if not group:
            continue
        if isinstance(group, str):
            group = [part.strip() for part in group.split(',')]
        for item in group:
            model = str(item or '').strip()
            if not model or model in seen:
                continue
            if any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._-' for c in model):
                continue
            seen.add(model)
            result.append(model)
    return result


def load_runtime_config() -> dict:
    primary = os.getenv('GEMINI_PRIMARY_MODEL', DEFAULT_PRIMARY_MODEL).strip() or DEFAULT_PRIMARY_MODEL
    fallback = normalize_model_list(os.getenv('GEMINI_FALLBACK_MODELS', ','.join(DEFAULT_FALLBACK_MODELS)))
    fallback = [m for m in fallback if m != primary]
    api_key = os.getenv('GEMINI_API_KEY', '').strip()
    return {
        'hasApiKey': bool(api_key),
        'apiKey': api_key,
        'primaryModel': primary,
        'fallbackModels': fallback,
        'allowBrowserOverride': False,
    }


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
            return res.status, res.read()
    except urllib.error.HTTPError as err:
        return err.code, err.read()


class Handler(SimpleHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        if self.path in ('/api/gemini-config', '/api/gemini-config/'):
            cfg = load_runtime_config()
            body = json.dumps({
                'hasApiKey': cfg['hasApiKey'],
                'primaryModel': cfg['primaryModel'],
                'fallbackModels': cfg['fallbackModels'],
                'allowBrowserOverride': False,
            }).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        # blok file sensitif bila nanti ada yang terunggah tidak sengaja
        sensitive = ('gemini.config.json', 'gemini.config.local.json', '.env')
        if any(self.path.endswith(name) for name in sensitive):
            self.send_error(404)
            return
        return super().do_GET()

    def do_POST(self):
        if self.path == '/api/export-word':
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if length <= 0 or length > 45 * 1024 * 1024:
                    raise ValueError('Ukuran data export Word tidak valid.')
                payload = json.loads(self.rfile.read(length).decode('utf-8'))
                pages = payload.get('pages')
                if not isinstance(pages, list) or not 1 <= len(pages) <= 6:
                    raise ValueError('Halaman Word tidak valid.')
                raw_name = str(payload.get('filename') or 'MCU_Medical_Checkup.docx')
                filename = re.sub(r'[^A-Za-z0-9._-]+', '_', raw_name).strip('._') or 'MCU_Medical_Checkup.docx'
                if not filename.lower().endswith('.docx'):
                    filename += '.docx'

                document = Document()
                section = document.sections[0]
                section.page_width = Mm(210)
                section.page_height = Mm(297)
                section.top_margin = Mm(0)
                section.bottom_margin = Mm(0)
                section.left_margin = Mm(0)
                section.right_margin = Mm(0)
                section.header_distance = Mm(0)
                section.footer_distance = Mm(0)

                for index, encoded in enumerate(pages):
                    if not isinstance(encoded, str):
                        raise ValueError('Format gambar halaman tidak valid.')
                    try:
                        image_bytes = base64.b64decode(encoded, validate=True)
                    except Exception as exc:
                        raise ValueError('Data gambar Word rusak.') from exc
                    if len(image_bytes) > 12 * 1024 * 1024:
                        raise ValueError('Satu halaman Word terlalu besar.')
                    paragraph = document.add_paragraph()
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1
                    run = paragraph.add_run()
                    run.add_picture(io.BytesIO(image_bytes), width=Mm(209), height=Mm(296))
                    if index < len(pages) - 1:
                        document.add_page_break()

                output = io.BytesIO()
                document.save(output)
                body = output.getvalue()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as exc:
                body = json.dumps({'error': {'message': str(exc)}}).encode('utf-8')
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            return

        if self.path != '/api/gemini':
            self.send_error(404)
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if length <= 0 or length > 80 * 1024 * 1024:
                raise ValueError('Ukuran permintaan tidak valid.')
            payload = json.loads(self.rfile.read(length).decode('utf-8'))
            request_body = payload.get('request')
            if not isinstance(request_body, dict):
                raise ValueError('Body permintaan Gemini tidak valid.')

            cfg = load_runtime_config()
            key = cfg['apiKey']
            if not key:
                raise ValueError('GEMINI_API_KEY belum diatur pada Railway Variables.')

            models = normalize_model_list(cfg['primaryModel'], cfg['fallbackModels'])
            if not models:
                raise ValueError('Model Gemini belum diatur.')

            attempts = []
            final_status, final_body = 500, b''
            for idx, model in enumerate(models):
                status, body = call_gemini(model, key, request_body)
                attempts.append({'model': model, 'status': status})
                if 200 <= status < 300:
                    self.send_response(status)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.send_header('Content-Length', str(len(body)))
                    self.send_header('X-Gemini-Model-Used', model)
                    self.send_header('X-Gemini-Attempts', json.dumps(attempts, ensure_ascii=False))
                    self.end_headers()
                    self.wfile.write(body)
                    return
                final_status, final_body = status, body
                if status in (400, 401, 403):
                    break
                if status not in RETRYABLE_STATUS and idx == len(models) - 1:
                    break

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
    host = os.getenv('HOST', '0.0.0.0')
    port = int(os.getenv('PORT', '8080'))
    server = ThreadingHTTPServer((host, port), Handler)
    print(f'Generator MCU Railway listening on http://{host}:{port}')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
