"""Simple HTTP server that binds to 0.0.0.0 for LAN access.

Usage:
    python app/serve.py
    python app/serve.py 8765
"""
import http.server
import os
import sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app' if os.path.basename(os.path.dirname(os.path.abspath(__file__))) != 'app' else '')

handler = http.server.SimpleHTTPRequestHandler
with http.server.HTTPServer(('0.0.0.0', PORT), handler) as httpd:
    print(f'Serving on http://0.0.0.0:{PORT}')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
