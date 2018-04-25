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
        self.send_response(200)
        self.send_header("Content-type", "text/xml")
        self.end_headers()
        response_string = """
<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Header>
    <t:ServerVersionInfo MajorVersion="8" MinorVersion="0" MajorBuildNumber="685" MinorBuildNumber="8"
                         xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
  </soap:Header>
  <soap:Body>
    <BogusResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <m:ResponseMessages>
        <m:BogusResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
        </m:BogusResponseMessage>
      </m:ResponseMessages>
    </BogusResponse>
  </soap:Body>
</soap:Envelope>
        """
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
