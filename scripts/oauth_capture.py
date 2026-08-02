import http.server, webbrowser, json, urllib.parse, sys, threading

CLIENT_ID = "1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com"
CLIENT_SECRET = "v6V3fKV_zWU7iw1DrpO1rknX"
PORT = 8888
REDIRECT = f"http://localhost:{PORT}"

SCOPES = [
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.webapp.deploy",
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/service.management",
    "https://www.googleapis.com/auth/logging.read",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/cloud-platform",
]

code_holder = {}

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        if "code" in qs:
            code_holder["code"] = qs["code"][0]
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(b"<h2>Authorization successful!</h2><p>You can close this tab now.</p>")
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"no code")
    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

server = http.server.HTTPServer(("localhost", PORT), Handler)
threading.Thread(target=server.serve_forever, daemon=True).start()

params = urllib.parse.urlencode({
    "redirect_uri": REDIRECT,
    "access_type": "offline",
    "scope": " ".join(SCOPES),
    "response_type": "code",
    "client_id": CLIENT_ID,
})
url = f"https://accounts.google.com/o/oauth2/v2/auth?{params}"
print("AUTH_URL", url)
webbrowser.open(url)

for _ in range(180):
    if "code" in code_holder:
        server.shutdown()
        print("CODE", code_holder["code"])
        sys.exit(0)
    import time; time.sleep(1)
print("TIMEOUT")
sys.exit(1)
