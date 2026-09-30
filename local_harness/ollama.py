"""Ollama's native local REST API, with streaming and interruptible sockets."""

import http.client
import json
import socket
import threading
from contextlib import contextmanager
from urllib.parse import urlsplit


class OllamaError(RuntimeError):
    pass


class Cancelled(OllamaError):
    pass


class OllamaClient:
    def __init__(self, base_url="http://127.0.0.1:11434"):
        try:
            parsed = urlsplit(base_url.strip())
            if (parsed.scheme != "http" or parsed.hostname not in {"localhost", "127.0.0.1", "::1"} or
                    parsed.username is not None or parsed.password is not None or
                    parsed.path not in {"", "/"} or parsed.query or parsed.fragment):
                raise ValueError()
            self.host = parsed.hostname
            self.port = parsed.port or 11434
        except (ValueError, AttributeError) as exc:
            raise OllamaError("Use a local Ollama address, such as http://127.0.0.1:11434 (no keys or URL paths).") from exc
        self.base_url = base_url.rstrip("/")
        self._lock = threading.Lock()
        self._connection = None
        self._response_socket = None
        self._response_reader = None
        self._cancelled = threading.Event()

    def cancel(self):
        self._cancelled.set()
        with self._lock:
            connection = self._connection
            sockets = [self._response_socket, connection.sock if connection else None]
            reader = self._response_reader
        # Shutdown releases a worker blocked on response headers or readline. No Tk calls here.
        for connection_socket in sockets:
            if connection_socket is not None:
                try:
                    connection_socket.shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass
        # Windows may keep recv_into blocked after shutdown while makefile owns an I/O reference.
        # Closing its raw reader (not the locked BufferedReader) releases that reference.
        for connection_socket in sockets:
            if connection_socket is not None:
                connection_socket.close()
        if reader is not None:
            reader.close()

    def _check_cancelled(self, cancelled):
        if self._cancelled.is_set() or (cancelled is not None and cancelled.is_set()):
            raise Cancelled("Stopped. No further tool calls will run.")

    @contextmanager
    def _exchange(self, path, body=None, timeout=15, cancelled=None):
        self._cancelled.clear()
        self._check_cancelled(cancelled)
        connection = http.client.HTTPConnection(self.host, self.port, timeout=timeout)
        response = None

        def make_response(sock, **kwargs):
            result = http.client.HTTPResponse(sock, **kwargs)
            with self._lock:
                self._response_reader = result.fp.raw
            return result

        connection.response_class = make_response
        with self._lock:
            self._connection = connection
        try:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
            connection.request("POST" if body is not None else "GET", path, body=data,
                               headers={"Content-Type": "application/json"})
            # Keep the socket before getresponse() can detach it for HTTP/1.0/Connection: close.
            with self._lock:
                self._response_socket = connection.sock
            self._check_cancelled(cancelled)
            response = connection.getresponse()
            self._check_cancelled(cancelled)
            if response.status >= 400:
                error = response.read(65536).decode("utf-8", "replace")
                try:
                    error = json.loads(error).get("error", error)
                except (ValueError, AttributeError):
                    pass
                if response.status == 404:
                    raise OllamaError(f"Model not installed or endpoint not found: {str(error)[:500]}. Refresh models or install it in Ollama.")
                raise OllamaError(f"Ollama HTTP {response.status}: {str(error)[:500]}")
            yield response
        except (OSError, ValueError, http.client.HTTPException) as exc:
            self._check_cancelled(cancelled)
            raise OllamaError("Cannot reach Ollama. Start Ollama (or run ollama serve), then Refresh models. "
                              "If generation timed out, try a smaller model/context.") from exc
        finally:
            with self._lock:
                self._connection = None
                self._response_socket = None
                self._response_reader = None
            if response is not None:
                try:
                    response.close()
                except (OSError, ValueError):
                    pass  # Stop may already have closed the raw reader.
            connection.close()

    def _json(self, path, body=None, cancelled=None):
        with self._exchange(path, body, cancelled=cancelled) as response:
            raw = response.read(4 * 1024 * 1024 + 1)
            if len(raw) > 4 * 1024 * 1024:
                raise OllamaError("Ollama returned an oversized metadata response.")
            try:
                value = json.loads(raw)
                if not isinstance(value, dict):
                    raise ValueError()
                return value
            except (ValueError, UnicodeError) as exc:
                raise OllamaError("Ollama returned invalid JSON.") from exc

    def list_models(self):
        return self._json("/api/tags").get("models", [])

    def show_model(self, model, cancelled=None):
        return self._json("/api/show", {"model": model}, cancelled=cancelled)

    def chat(self, model, messages, tools=None, options=None, think=False, on_chunk=None, cancelled=None):
        payload = {"model": model, "messages": messages, "stream": True,
                   "options": options or {}, "keep_alive": "60s", "think": think}
        if tools:
            payload["tools"] = tools
        result = {"role": "assistant", "content": "", "thinking": "", "tool_calls": [], "stats": {}}
        done = False
        # This is an idle-read timeout, not a short total-turn deadline. CPU prompt evaluation can be slow.
        with self._exchange("/api/chat", payload, timeout=1200, cancelled=cancelled) as response:
            while True:
                self._check_cancelled(cancelled)
                line = response.readline(1024 * 1024 + 1)
                self._check_cancelled(cancelled)
                if not line:
                    break
                if len(line) > 1024 * 1024:
                    raise OllamaError("Ollama returned an oversized stream chunk.")
                if not line.strip():
                    continue
                try:
                    chunk = json.loads(line)
                    if not isinstance(chunk, dict):
                        raise ValueError()
                except (ValueError, UnicodeError) as exc:
                    raise OllamaError("Ollama returned an invalid streaming response.") from exc
                if chunk.get("error"):
                    raise OllamaError(str(chunk["error"])[:500])
                message = chunk.get("message") or {}
                if not isinstance(message, dict):
                    raise OllamaError("Invalid message from Ollama.")
                for key in ("content", "thinking"):
                    value = message.get(key) or ""
                    if not isinstance(value, str):
                        raise OllamaError("Invalid text from Ollama.")
                    result[key] += value
                calls = message.get("tool_calls") or []
                if not isinstance(calls, list):
                    raise OllamaError("Invalid tool calls from Ollama.")
                result["tool_calls"].extend(calls)
                if len(result["content"]) + len(result["thinking"]) > 256000:
                    raise OllamaError("Generation exceeded the harness output limit.")
                if on_chunk:
                    on_chunk(message)
                if chunk.get("done"):
                    done = True
                    result["stats"] = {k: v for k, v in chunk.items() if k not in {"message", "model", "created_at"}}
                    break
        if not done:
            raise OllamaError("Ollama closed the stream before completion. Retry the turn.")
        return result
