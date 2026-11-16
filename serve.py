"""Serves this folder on localhost:8080 with caching off, so edits to config.js show on reload.
POST /imap is the bridge to Yahoo Mail (imap_bridge.py)."""
import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

import imap_bridge

PAGE_ORIGINS = {'http://localhost:8080', 'http://127.0.0.1:8080'}


class NoCache(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self):
        # Only this app's own page may use the bridge: other sites' pages send their own Origin. No CORS headers on purpose.
        host_ok = self.headers.get('Host') in {o.split('//')[1] for o in PAGE_ORIGINS}
        if self.path != '/imap' or self.headers.get('Origin') not in PAGE_ORIGINS or not host_ok:
            self.send_error(403)
            return
        try:
            request = json.loads(self.rfile.read(int(self.headers.get('Content-Length', 0))))
            status, body = imap_bridge.handle(request)
        except ValueError:
            status, body = 400, {'error': 'Bad request.'}
        data = json.dumps(body).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)


if __name__ == '__main__':
    ThreadingHTTPServer(("127.0.0.1", 8080), NoCache).serve_forever()
