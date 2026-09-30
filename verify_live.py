"""Optional live check against your local Ollama. Does not install or remove models."""

import argparse
import json
import tempfile
import time
from pathlib import Path

from local_harness.agent import Agent, Attachment
from local_harness.ollama import OllamaClient, OllamaError
from local_harness.workspace import Workspace


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default="qwen3:0.6b")
    parser.add_argument("--mode", choices=("auto", "native", "json"), default="auto")
    args = parser.parse_args()
    client = OllamaClient()
    names = [m["name"] for m in client.list_models()]
    print("Installed models:", ", ".join(names), flush=True)
    for name in (f"mimo-2.6-{level}{suffix}" for level in ("low", "medium", "high") for suffix in ("-fast", "")):
        if name not in names and name + ":latest" not in names:
            try:
                client.show_model(name)
            except OllamaError:
                print(name + ": missing model correctly rejected (no fallback)", flush=True)
            else:
                raise AssertionError("Expected a missing model error")
    with tempfile.TemporaryDirectory(prefix="local-harness-") as temp:
        root = Path(temp)
        (root / "notes.txt").write_text("Verification code: HARNESS-4821\n", encoding="utf-8")
        events = []

        def event(kind, value):
            events.append((kind, value))
            if kind == "tool":
                print("Tool:", value["name"], value.get("path", ""), flush=True)
            elif kind == "status":
                print(value, flush=True)

        agent = Agent(client, args.model, Workspace(root), tool_mode=args.mode,
                      options={"num_ctx": 4096, "num_predict": 192, "temperature": 0}, on_event=event)
        started = time.monotonic()
        answer = agent.run("Call read_file with path notes.txt, then reply with only the verification code found in that file.")
        print("Answer:", answer, flush=True)
        assert any(k == "tool_result" and v["name"] == "read_file" and "HARNESS-4821" in str(v["result"])
                   for k, v in events), "The model must actually call read_file, not pretend."
        assert "HARNESS-4821" in answer, "The final answer must use the real tool result."
        print("PASS: live file tool + exact model selection; seconds:", round(time.monotonic() - started, 1), flush=True)
        attached = Agent(client, args.model, options={"num_ctx": 2048, "num_predict": 64, "temperature": 0})
        answer = attached.run("Quote the complete Code value from the attached file, keeping all letters, hyphens and digits.",
                              [Attachment("attachment.txt", "Code: ATTACH-7319")])
        print("Attachment answer:", answer, flush=True)
        assert "ATTACH-7319" in answer, "Attachment content must reach the selected model."
        print("PASS: attachment-only chat:", answer, flush=True)


if __name__ == "__main__":
    main()
