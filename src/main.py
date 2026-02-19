from __future__ import annotations

import tkinter as tk

from .gui import ExcelParserApp


def main() -> None:
    root = tk.Tk()
    ExcelParserApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

