import json
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from local_harness.ollama import Cancelled, OllamaClient, OllamaError


class TestServer(ThreadingHTTPServer):
    daemon_threads = True


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'{"models":[{"name":"qwen-test:latest","size":1234}]}')

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        self.server.requests.append((self.path, body))
        if body["model"] == "mimo-missing":
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'{"error":"model mimo-missing not found"}')
            return
        self.send_response(200)
        self.end_headers()
        if body["model"] == "stall":
            self.wfile.write(b'{"message":{"content":"Started"},"done":false}\n')
            self.wfile.flush()
            self.server.started.set()
            self.server.release.wait(3)
            return
        if self.path == "/api/show":
            self.wfile.write(b'{"capabilities":["completion","tools"]}')
        else:
            chunks = [
                {"message": {"content": "Hello "}, "done": False},
                {"message": {"content": "world"}, "done": False},
                {"message": {"tool_calls": [{"function": {"name": "read_file", "arguments": {"path": "a.txt"}}}]}, "done": False},
                {"message": {}, "done": True, "eval_count": 2, "eval_duration": 1000000000},
            ]
            for chunk in chunks:
                self.wfile.write(json.dumps(chunk).encode() + b"\n")
                self.wfile.flush()


class OllamaTests(unittest.TestCase):
    def setUp(self):
        self.server = TestServer(("127.0.0.1", 0), Handler)
        self.server.requests = []
        self.server.started = threading.Event()
        self.server.release = threading.Event()
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.client = OllamaClient(f"http://127.0.0.1:{self.server.server_port}")

    def tearDown(self):
        self.client.cancel()
        self.server.release.set()
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=2)

    def test_discover_select_and_stream_the_exact_model_with_tools(self):
        self.assertEqual(self.client.list_models()[0]["name"], "qwen-test:latest")
        self.assertIn("tools", self.client.show_model("qwen-test:latest")["capabilities"])
        pieces = []
        reply = self.client.chat(
            "qwen-test:latest", [{"role": "user", "content": "hello"}],
            on_chunk=lambda chunk: pieces.append(chunk), options={"num_ctx": 4096},
        )
        self.assertEqual(reply["content"], "Hello world")
        self.assertEqual(reply["tool_calls"][0]["function"]["name"], "read_file")
        self.assertEqual(reply["stats"]["eval_count"], 2)
        self.assertEqual(self.server.requests[-1][1]["model"], "qwen-test:latest")
        self.assertEqual(self.server.requests[-1][1]["options"]["num_ctx"], 4096)
        self.assertGreaterEqual(len(pieces), 3)

    def test_missing_mimo_is_an_error_not_a_fallback_to_another_model(self):
        with self.assertRaises(OllamaError):
            self.client.show_model("mimo-missing")
        self.assertEqual(self.server.requests[-1][1]["model"], "mimo-missing")

    def test_rejects_cloud_addresses_and_embedded_credentials(self):
        for url in ("https://example.com", "http://user:password@localhost:11434", "http://localhost:11434/api"):
            with self.assertRaises(OllamaError):
                OllamaClient(url)

    def test_stop_interrupts_a_stalled_stream(self):
        errors = []

        def run():
            try:
                self.client.chat("stall", [{"role": "user", "content": "hello"}])
            except Exception as exc:
                errors.append(exc)

        worker = threading.Thread(target=run, daemon=True)
        worker.start()
        self.assertTrue(self.server.started.wait(2))
        self.client.cancel()
        worker.join(timeout=2)
        self.assertFalse(worker.is_alive(), "Stop must not wait for the 20-minute read timeout.")
        self.assertEqual(len(errors), 1)
        self.assertIsInstance(errors[0], Cancelled)


if __name__ == "__main__":
    unittest.main()
