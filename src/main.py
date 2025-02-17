import logging
import os
import sys
import tkinter as tk
from tkinter import filedialog

from update_fields import update_fields

logging.basicConfig()
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

FILENAME = (
    "..\\samples\\input.docx" if sys.platform == "win32" else "../samples/input.docx"
)


def _select_file() -> str:
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()
    return file_path


def main() -> None:
    source_file: str = os.path.abspath(FILENAME)

    if not os.path.exists(source_file):
        source_file = _select_file()

    update_fields(source_file)


if __name__ == "__main__":
    main()
