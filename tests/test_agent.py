import tempfile
import unittest
from pathlib import Path

from local_harness.agent import Agent, Attachment
from local_harness.workspace import Workspace
from local_harness.ollama import OllamaError


class ModelAdapter:
    """A deterministic external-model substitute, not an internal implementation mock."""
    def __init__(self, replies, capabilities=("tools",)):
        self.replies = iter(replies)
        self.capabilities = capabilities
        self.requests = []

    def show_model(self, model, cancelled=None):
        return {"capabilities": list(self.capabilities)}

    def chat(self, model, messages, **kwargs):
        self.requests.append((model, [dict(m) for m in messages], kwargs))
        reply = next(self.replies)
        if kwargs.get("on_chunk"):
            kwargs["on_chunk"]({"content": reply.get("content", "")})
        return {"role": "assistant", "thinking": "", "tool_calls": [], "stats": {}, **reply}

    def cancel(self):
        pass


class AgentTests(unittest.TestCase):
    def test_native_file_call_returns_real_contents_to_the_selected_model(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "sales.csv").write_text("item,total\napples,17\n", encoding="utf-8")
            adapter = ModelAdapter([
                {"tool_calls": [{"function": {"name": "read_file", "arguments": {"path": "sales.csv"}}}]},
                {"content": "The apples total is 17."},
            ])
            agent = Agent(adapter, "qwen-test", Workspace(root))
            answer = agent.run("Read sales.csv")
            self.assertEqual(answer, "The apples total is 17.")
            result = adapter.requests[-1][1][-1]
            self.assertEqual(result["role"], "tool")
            self.assertEqual(result["tool_name"], "read_file")
            self.assertIn("apples,17", result["content"])
            self.assertTrue(all(r[0] == "qwen-test" for r in adapter.requests))

    def test_attachments_are_supplied_without_granting_folder_access(self):
        adapter = ModelAdapter([{ "content": "Noted." }])
        agent = Agent(adapter, "qwen-test")
        self.assertEqual(agent.run("Review this", [Attachment("notes.txt", "attached fact")]), "Noted.")
        self.assertIn("attached fact", adapter.requests[0][1][-1]["content"])
        self.assertIsNone(adapter.requests[0][2].get("tools"))

    def test_json_compatibility_executes_tools_without_native_tool_support(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "notes.txt").write_text("ground truth", encoding="utf-8")
            adapter = ModelAdapter([
                {"content": '```json\n{"tool":"read_file","arguments":{"path":"notes.txt"}}\n```'},
                {"content": "ground truth"},
            ], capabilities=())
            answer = Agent(adapter, "mimo-test", Workspace(root)).run("Read notes.txt")
            self.assertEqual(answer, "ground truth")
            self.assertIsNone(adapter.requests[0][2]["tools"])
            self.assertIn('"tool"', adapter.requests[0][1][0]["content"])
            self.assertIn("ground truth", adapter.requests[1][1][-1]["content"])

    def test_mimo_legacy_xml_tool_call_is_understood(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "notes.txt").write_text("42", encoding="utf-8")
            adapter = ModelAdapter([
                {"content": '<tool_call><function=read_file>{"path":"notes.txt"}</function></tool_call>'},
                {"content": "42"},
            ])
            self.assertEqual(Agent(adapter, "mimo-test", Workspace(root)).run("Read notes.txt"), "42")

    def test_write_denial_is_reported_to_the_model_without_changing_the_file(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "notes.txt").write_text("keep", encoding="utf-8")
            adapter = ModelAdapter([
                {"tool_calls": [{"function": {"name": "write_file", "arguments": {"path": "notes.txt", "content": "replace"}}}]},
                {"content": "The write was denied."},
            ])
            Agent(adapter, "qwen-test", Workspace(root), approve=lambda request: False).run("Replace notes.txt")
            self.assertEqual((root / "notes.txt").read_text(encoding="utf-8"), "keep")
            self.assertIn('"status": "denied"', adapter.requests[-1][1][-1]["content"])

    def test_oversized_input_is_rejected_before_inference(self):
        adapter = ModelAdapter([{"content": "should not be generated"}])
        with self.assertRaises(OllamaError):
            Agent(adapter, "qwen-test").run("x" * 50000)
        self.assertEqual(adapter.requests, [])

    def test_json_embedded_in_explanatory_prose_is_not_executed(self):
        with tempfile.TemporaryDirectory() as temp:
            reply = 'Example: {"tool":"write_file","arguments":{"path":"a.txt","content":"no"}}'
            adapter = ModelAdapter([{"content": reply}])
            self.assertEqual(Agent(adapter, "qwen-test", Workspace(temp)).run("Explain tool syntax"), reply)
            self.assertFalse((Path(temp) / "a.txt").exists())

    def test_plain_answers_stream_before_completion_in_json_mode(self):
        observed = []
        text_events = []

        class StreamingAdapter(ModelAdapter):
            def chat(self, model, messages, **kwargs):
                kwargs["on_chunk"]({"content": "Hel"})
                kwargs["on_chunk"]({"content": "lo"})
                observed.append("".join(text_events))
                return {"role": "assistant", "content": "Hello", "tool_calls": [], "stats": {}}

        with tempfile.TemporaryDirectory() as temp:
            agent = Agent(StreamingAdapter([], capabilities=()), "test", Workspace(temp), tool_mode="json",
                          on_event=lambda kind, value: text_events.append(value) if kind == "text" else None)
            self.assertEqual(agent.run("Say hello"), "Hello")
        self.assertEqual(observed, ["Hello"])
        self.assertEqual("".join(text_events), "Hello", "Streaming must not duplicate the final answer.")


if __name__ == "__main__":
    unittest.main()
