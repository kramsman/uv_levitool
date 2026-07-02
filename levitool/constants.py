"""Shared constants for the LeviTool package."""

from pathlib import Path

INITIAL_ATTACHMENT_DIR = Path("~/Dropbox/Postcard Files/Attachments/Campaigns/").expanduser()

# Keys required in every BOE xlsx, regardless of docx template contents
PROGRAM_REQUIRED_KEYS = {'{STATE}', '{CNTYFILENAME}', '{USE}'}

# --- TEST FLAGS: set True only during dev, never commit as True ---
TEST_INPUT_DIR = False       # skip file picker, use hardcoded input path
TEST_SKIP_PROMPTS = False    # skip exit_yes_no confirmation dialogs

# Assumed font size (pt) when a label line specifies no readable Latin size (w:sz).
# Matches Word's built-in fallback when nothing in the file specifies a size.
DEFAULT_FONT_SIZE_PT = 10

# Lower-end characters per line for 30 per page Avery 5161 label, keyed by font size (pt)
CHARS_PER_LINE_30PP = {
    8: 52,
    9: 46,
    10: 42,
    11: 38,
    12: 35,
    13: 32,
    14: 30,
}

MAX_CHARS_PER_LINE_TEXT = """Estimated character limits

30 per page, Avery 5161 Label
Font Size | Approximate Characters per Line
----------|--------------------------------
8 pt        52-55 characters
9 pt        46-49 characters
10 pt       42-45 characters
11 pt       38-41 characters
12 pt       35-38 characters
13 pt       32-35 characters
14 pt       30-33 characters

20 per page, Avery 5161 Label
Font Size | Approximate Characters per Line
----------|--------------------------------
8 pt        80-85 characters
9 pt        71-76 characters
10 pt       64-68 characters
11 pt       58-62 characters
12 pt       53-57 characters
13 pt       49-52 characters
14 pt       46-49 characters"""
