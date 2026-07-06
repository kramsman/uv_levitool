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
# NOTE: the old CHARS_PER_LINE_30PP / MAX_CHARS_PER_LINE_TEXT character-count tables
# were removed; line fit is now measured from the real font width (see label_width.py).
DEFAULT_FONT_SIZE_PT = 10
