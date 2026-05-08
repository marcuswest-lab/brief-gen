#!/usr/bin/env python3
"""Local dev server for the brief generator.

Pinned to an explicit directory because macOS sandboxing can deny os.getcwd()
on the parent path, which breaks SimpleHTTPRequestHandler's default behavior.
"""

import os
import functools
import http.server
import socketserver

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = int(os.environ.get('PORT', '8766'))


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        '.js':   'application/javascript',
        '.mjs':  'application/javascript',
        '.json': 'application/json',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }

    def end_headers(self):
        # Disable caching during development so edits show up immediately
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()


def main():
    handler = functools.partial(Handler, directory=ROOT)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(('127.0.0.1', PORT), handler) as httpd:
        print(f'Serving {ROOT} on http://127.0.0.1:{PORT}')
        httpd.serve_forever()


if __name__ == '__main__':
    main()
