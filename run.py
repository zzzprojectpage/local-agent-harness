"""Launch the native desktop app. Python 3.10+ with Tk; Setup.cmd adds Excel tools."""

import tkinter as tk
from local_harness.app import HarnessApp


def main():
    root = tk.Tk()
    HarnessApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
