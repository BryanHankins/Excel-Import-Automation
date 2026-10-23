"""The default engine talks to a local Ollama server; these tests use a fake one."""
import io
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest
from PIL import Image

from drugtest_import import extract
from drugtest_import.forms import DRUG_TEST


class FakeOllama(BaseHTTPRequestHandler):
    answer = {}
    status = 200
    error = ""
    models = ["qwen3-vl:8b"]
    requests = []

    def log_message(self, *args):
        pass

    def _send(self, status, payload):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self._send(200, {"models": [{"name": n} for n in FakeOllama.models]})

    def do_POST(self):
        request = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        FakeOllama.requests.append(request)
        if FakeOllama.status != 200:
            return self._send(FakeOllama.status, {"error": FakeOllama.error})
        content = FakeOllama.answer if isinstance(FakeOllama.answer, str) else json.dumps(FakeOllama.answer)
        self._send(200, {"model": request["model"], "message": {"role": "assistant", "content": content},
                         "prompt_eval_count": 900, "eval_count": 120, "done": True})


@pytest.fixture
def ollama(monkeypatch):
    server = HTTPServer(("127.0.0.1", 0), FakeOllama)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    monkeypatch.setattr(extract, "ENGINE", "ollama")
    monkeypatch.setattr(extract, "MODEL", "qwen3-vl:8b")
    monkeypatch.setattr(extract, "OLLAMA_URL", f"http://127.0.0.1:{server.server_port}")
    FakeOllama.status, FakeOllama.requests = 200, []
    yield FakeOllama
    server.shutdown()


def image():
    buffer = io.BytesIO()
    Image.new("RGB", (2400, 1600), "yellow").save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def test_reads_form_with_local_model(ollama):
    ollama.answer = {"employee_id": "emp005", "name": " Bob Wilson ", "department": "", "test_date": "9/18/2025",
                     "test_type": "urine", "result": "Negative", "notes": "null",
                     "uncertain_fields": ["result", "not_a_field"]}
    reading = extract.read_form(image(), DRUG_TEST)
    assert reading.values == {"employee_id": "emp005", "name": "Bob Wilson", "department": None,
                              "test_date": "9/18/2025", "test_type": "urine", "result": "Negative", "notes": None}
    assert reading.uncertain == ["result"]
    assert (reading.model, reading.input_tokens, reading.output_tokens) == ("qwen3-vl:8b", 900, 120)

    [request] = ollama.requests
    assert request["model"] == "qwen3-vl:8b" and request["stream"] is False
    assert set(request["format"]["properties"]) == set(DRUG_TEST.keys) | {"uncertain_fields"}
    assert request["format"]["properties"]["uncertain_fields"]["items"]["enum"] == DRUG_TEST.keys
    message = request["messages"][0]
    assert "test_type: Test type" in message["content"]
    with Image.open(io.BytesIO(__import__("base64").b64decode(message["images"][0]))) as sent:
        assert sent.format == "JPEG" and max(sent.size) == extract.MAX_SIDE


def test_invalid_json_answer(ollama):
    ollama.answer = "not json"
    with pytest.raises(extract.ExtractionError, match="valid JSON"):
        extract.read_form(image(), DRUG_TEST)


def test_missing_model_explains_how_to_download(ollama):
    ollama.status, ollama.error = 404, 'model "qwen3-vl:8b" not found, try pulling it first'
    with pytest.raises(extract.ExtractionError, match="ollama pull qwen3-vl:8b"):
        extract.read_form(image(), DRUG_TEST)


def test_ollama_not_running(monkeypatch):
    monkeypatch.setattr(extract, "ENGINE", "ollama")
    monkeypatch.setattr(extract, "OLLAMA_URL", "http://127.0.0.1:9")  # nothing listens here
    with pytest.raises(extract.ExtractionError, match="Can't reach Ollama"):
        extract.read_form(image(), DRUG_TEST)


def test_check_ollama(ollama, monkeypatch):
    assert "ready" in extract.check_ollama()
    monkeypatch.setattr(extract, "MODEL", "gemma3:12b")
    with pytest.raises(extract.ExtractionError, match="ollama pull gemma3:12b"):
        extract.check_ollama()


def test_unknown_engine(monkeypatch):
    monkeypatch.setattr(extract, "ENGINE", "magic")
    with pytest.raises(extract.ExtractionError, match="Unknown DRUGTEST_ENGINE"):
        extract.read_form(image(), DRUG_TEST)
