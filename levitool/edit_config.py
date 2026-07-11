"""GUI editor for LeviTool's editable constants in constants.py.

Presents each managed constant with a short description of its purpose. Constants
whose values are limited to a fixed set (e.g. the PDF converter, the True/False test
flags) are shown as pick-lists; free-value constants (e.g. a font size) use a spin box
or text field. Saving writes the chosen values straight back into constants.py,
preserving the surrounding comments. Modelled on 4thU/.../pco_webhook_root/edit_config.py.
"""

import ast
import re
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QComboBox, QSpinBox,
    QLineEdit, QLabel, QFrame, QPushButton, QMessageBox,
)

CONSTANTS_FILE = Path(__file__).parent / "constants.py"

# Schema of the constants exposed in the editor. Each entry:
#   name        - the constant's identifier in constants.py
#   label       - short human label shown next to its control
#   type        - "choice" (pick from choices), "bool" (pick True/False),
#                 "int" (spin box), or "text" (free text)
#   choices     - allowed values for "choice"
#   description - one/two sentences on what the constant does
CONFIG_FIELDS = [
    {
        "name": "DOC_TO_PDF_CONVERTER",
        "label": "PDF converter",
        "type": "choice",
        "choices": ["docx2pdf", "wordexporter", "libreoffice"],
        "description": "How the merged .docx files are turned into PDF. "
                       "'docx2pdf' uses Word's Save-As-PDF exporter (fast, but drops the color CFCG "
                       "symbol). 'wordexporter' drives Word's Print → Save as PDF (keeps the color symbol; "
                       "slower, needs one-time macOS Automation/Accessibility permission). 'libreoffice' "
                       "converts headless via LibreOffice (keeps the color symbol, no template changes; "
                       "LibreOffice must be installed).",
    },
    {
        "name": "DEFAULT_FONT_SIZE_PT",
        "label": "Default font size (pt)",
        "type": "int",
        "min": 4,
        "max": 72,
        "description": "Point size assumed when a label line specifies no readable Latin font size. "
                       "Matches Word's built-in fallback and is used by the label line-length fit check.",
    },
    {
        "name": "TEST_INPUT_DIR",
        "label": "TEST: skip input picker",
        "type": "bool",
        "description": "Developer flag. When True, skips the input-directory file picker and uses a "
                       "hardcoded path. Leave False for normal use.",
    },
    {
        "name": "TEST_SKIP_PROMPTS",
        "label": "TEST: skip confirmations",
        "type": "bool",
        "description": "Developer flag. When True, skips the yes/no confirmation dialogs. "
                       "Leave False for normal use.",
    },
    # --- Structural constants (kept at the bottom): a folder path and a set of keys. ---
    {
        "name": "INITIAL_ATTACHMENT_DIR",
        "label": "Initial campaigns folder",
        "type": "path",
        "description": "Starting folder shown when picking a campaign's input directory. Point it at your "
                       "Campaigns folder. Use ~ for your home directory (it is expanded automatically).",
    },
    {
        "name": "PROGRAM_REQUIRED_KEYS",
        "label": "Required BOE keys",
        "type": "keys",
        "description": "Merge keys that must exist in every BOE xlsx regardless of the docx templates "
                       "(e.g. {STATE}, {CNTYFILENAME}, {USE}). Enter as a comma-separated list.",
    },
]


def _read_current_value(text: str, field: dict):
    """Reads a constant's current literal from the constants.py text and parses it.

    Returns the parsed value (str/int/bool), or None if the constant is not found.
    """
    match = _assignment_re(field["name"]).search(text)
    if match is None:
        return None
    raw = match.group("val").strip()
    ftype = field["type"]
    if ftype == "bool":
        return raw == "True"
    if ftype == "int":
        try:
            return int(raw)
        except ValueError:
            return None
    if ftype == "path":
        # e.g. Path("~/.../Campaigns/").expanduser() -> show just the quoted path
        inner = re.search(r'''Path\(\s*["']([^"']*)["']''', raw)
        return inner.group(1) if inner else raw
    if ftype == "keys":
        # e.g. {'{STATE}', '{USE}'} -> show as a comma-separated list
        try:
            return ", ".join(sorted(ast.literal_eval(raw)))
        except (ValueError, SyntaxError):
            return raw
    # choice / text: strip surrounding quotes
    return raw.strip('"').strip("'")


def _assignment_re(name: str) -> re.Pattern:
    """Regex matching a top-level ``NAME = <value>[  # comment]`` line in constants.py."""
    return re.compile(
        rf'^(?P<pre>[ \t]*{re.escape(name)}[ \t]*=[ \t]*)'
        rf'(?P<val>.*?)(?P<cmt>[ \t]*#.*)?$',
        re.MULTILINE,
    )


def _literal(field: dict, value) -> str:
    """Formats a Python source literal for a field's new value."""
    ftype = field["type"]
    if ftype == "bool":
        return "True" if value else "False"
    if ftype == "int":
        return str(int(value))
    if ftype == "path":
        return f'Path("{value}").expanduser()'
    if ftype == "keys":
        items = [k.strip() for k in str(value).split(",") if k.strip()]
        return "{" + ", ".join(f"'{k}'" for k in items) + "}"
    return f'"{value}"'   # choice / text


def _write_values(new_values: dict) -> None:
    """Writes new_values back into constants.py, preserving comments and layout.

    Args:
        new_values: {constant_name: parsed_value} for the managed constants.

    Raises:
        ValueError: If a managed constant's assignment line cannot be found.
    """
    text = CONSTANTS_FILE.read_text(encoding="utf-8")
    for field in CONFIG_FIELDS:
        name = field["name"]
        if name not in new_values:
            continue
        literal = _literal(field, new_values[name])

        def repl(m, _lit=literal):
            return f"{m.group('pre')}{_lit}{m.group('cmt') or ''}"

        text, n = _assignment_re(name).subn(repl, text, count=1)
        if n == 0:
            raise ValueError(f"Could not find '{name} = ...' in {CONSTANTS_FILE.name}")
    CONSTANTS_FILE.write_text(text, encoding="utf-8")


class ConfigEditor(QDialog):
    """Modal dialog listing each managed constant with a control and description."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("LeviTool — Edit Config")
        self.setMinimumWidth(640)
        text = CONSTANTS_FILE.read_text(encoding="utf-8")
        layout = QVBoxLayout(self)

        intro = QLabel(f"Editing {CONSTANTS_FILE.name}. Changes take effect the next time LeviTool is run.")
        intro.setWordWrap(True)
        intro.setStyleSheet("color: #444; padding: 2px 2px 8px 2px;")
        layout.addWidget(intro)

        self._widgets = {}
        for field in CONFIG_FIELDS:
            current = _read_current_value(text, field)

            row = QHBoxLayout()
            label = QLabel(field["label"] + ":")
            label.setMinimumWidth(190)
            row.addWidget(label)
            widget = self._make_widget(field, current)
            self._widgets[field["name"]] = widget
            row.addWidget(widget, 1)
            layout.addLayout(row)

            desc = QLabel(field["description"])
            desc.setWordWrap(True)
            desc.setStyleSheet("color: #666; padding: 0 0 10px 194px;")
            layout.addWidget(desc)

            line = QFrame()
            line.setFrameShape(QFrame.HLine)
            line.setStyleSheet("color: #ddd;")
            layout.addWidget(line)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._save)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _make_widget(self, field: dict, current):
        """Builds the control for one field, pre-set to its current value."""
        ftype = field["type"]
        if ftype in ("choice", "bool"):
            widget = QComboBox()
            choices = field["choices"] if ftype == "choice" else ["False", "True"]
            widget.addItems(choices)
            current_str = ("True" if current else "False") if ftype == "bool" else (current or "")
            idx = widget.findText(current_str)
            if idx < 0 and current_str:
                # value in the file isn't one of the known choices — surface it
                widget.insertItem(0, f"⚠ {current_str}")
                idx = 0
            widget.setCurrentIndex(max(idx, 0))
            return widget
        if ftype == "int":
            widget = QSpinBox()
            widget.setRange(field.get("min", 0), field.get("max", 1000))
            widget.setValue(int(current) if current is not None else 0)
            return widget
        widget = QLineEdit(str(current) if current is not None else "")
        return widget

    def _value_of(self, field: dict):
        """Reads the chosen value from a field's control, typed appropriately."""
        widget = self._widgets[field["name"]]
        ftype = field["type"]
        if ftype == "bool":
            return widget.currentText() == "True"
        if ftype == "choice":
            return widget.currentText().lstrip("⚠ ").strip()
        if ftype == "int":
            return widget.value()
        return widget.text().strip()

    def _save(self):
        new_values = {f["name"]: self._value_of(f) for f in CONFIG_FIELDS}
        try:
            _write_values(new_values)
        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not write {CONSTANTS_FILE.name}:\n{e}")
            return
        QMessageBox.information(
            self, "Saved",
            f"Saved to {CONSTANTS_FILE.name}.\nChanges take effect the next time you run LeviTool.")
        self.accept()


def edit_config() -> None:
    """Opens the constants editor. Safe to call from within LeviTool's menu."""
    app = QApplication.instance() or QApplication(sys.argv)
    dlg = ConfigEditor()
    dlg.show()
    dlg.raise_()
    dlg.activateWindow()
    dlg.exec()


if __name__ == "__main__":
    edit_config()
