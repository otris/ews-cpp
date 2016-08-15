from http.server import CGIHTTPRequestHandler, HTTPServer
import time
import sys

sys.stdout = None
sys.stderr = None

hostName = "localhost"
hostPort = 8080

class MyServer(CGIHTTPRequestHandler):
	def do_GET(self):
		time.sleep(2)
		self.send_response(200)
		self.send_header("Content-type", "text/html")
		self.end_headers()
		self.wfile.write(bytes("<html><head><title>hello</title></head>", "utf-8"))
	
	def do_POST(self):
		self.do_GET()

	def log_message(self, format, *args):
        	return

myServer = HTTPServer((hostName, hostPort), MyServer)
myServer.serve_forever()
