"""Live model-to-native-Excel verification using synthetic data, not private workbooks."""

import argparse
import hashlib
import tempfile
import zipfile
from pathlib import Path

import openpyxl

from local_harness.agent import Agent
from local_harness.ollama import OllamaClient
from local_harness.workspace import Workspace


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default="qwen3:0.6b")
    parser.add_argument("--mode", choices=("auto", "native", "json"), default="auto")
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="harness-live-excel-") as temp:
        root = Path(temp)
        book = openpyxl.Workbook()
        book.active.title = "Data"
        for row in (("Category", "Amount"), ("A", 10), ("A", 20), ("B", 5)):
            book.active.append(row)
        book.save(root / "data.xlsx")
        book.close()
        original = hashlib.sha256((root / "data.xlsx").read_bytes()).hexdigest()
        workspace = Workspace(root, autonomous=True)
        events = []

        def event(kind, value):
            events.append((kind, value))
            if kind == "tool":
                print("Tool:", value["name"], flush=True)
            elif kind in ("status", "diagnostic"):
                print(kind + ":", value, flush=True)
            elif kind == "tool_result" and isinstance(value["result"], dict):
                print("Result:", value["result"], flush=True)

        agent = Agent(OllamaClient(), args.model, workspace, tool_mode=args.mode,
                      options={"num_ctx": 8192, "num_predict": 768, "temperature": 0}, on_event=event)
        answer = agent.run('Create a native Excel PivotTable in data.xlsx, source Data!A1:B4, row field Category, '
                           'data field Amount with sum, target sheet Summary, target cell A3, name CategoryTotals. '
                           'Use create_pivot to perform the operation, then read_excel to verify the actual saved totals. '
                           'Do not just give instructions or example code.')
        print("Model answer:", answer, flush=True)
        writes = [v["result"] for k, v in events if k == "tool_result" and v["name"] == "create_pivot"
                  and v["result"].get("status") == "written"]
        assert writes, "The selected model did not execute a successful native pivot tool."
        with zipfile.ZipFile(root / writes[-1]["output_path"]) as archive:
            assert "xl/pivotTables/pivotTable1.xml" in archive.namelist()
        values = workspace.read_excel("data.xlsx", "Summary", "A3:B7")["values"]
        for expected in (["A", 30], ["B", 5], ["Grand Total", 35]):
            assert expected in values, (expected, values)
        assert hashlib.sha256((root / "data.xlsx").read_bytes()).hexdigest() == original
        print("PASS: selected local model created a native PivotTable; independently verified 30, 5, 35; original unchanged.", flush=True)


if __name__ == "__main__":
    main()
