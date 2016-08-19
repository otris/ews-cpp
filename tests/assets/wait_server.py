# -*- coding: utf-8 -*-

"""Launch small HTTP server for TimeoutTest test case

Should work with Python 2 and 3.

"""

import sys
import time

try:
    from SimpleHTTPServer import SimpleHTTPRequestHandler as RequestHandler
except ImportError:
    from http.server import CGIHTTPRequestHandler as RequestHandler

try:
    from SocketServer import TCPServer as HTTPServer
except ImportError:
    from http.server import HTTPServer


PYTHON_VERSION = sys.version_info[0]

class Handler(RequestHandler):
    def do_GET(self):
        time.sleep(2)
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.end_headers()
        response_string = "<html><head><title>hello</title></head>"
        if PYTHON_VERSION is 3:
            response = bytes(response_string, "utf-8")
        else:
            response = response_string
        self.wfile.write(response)

    def do_POST(self):
        self.do_GET()

    def log_message(self, format, *args):
        return


server = HTTPServer(("localhost", 8080), Handler)
server.serve_forever()
