# LeviTool

Create script and label files from Word (.docx) templates and an Excel (.xlsx) input file. Essentially a mail-merge tool designed for postcard writing campaigns.

## Features

- **LeviTool (Mail Merge)** - Reads county/district data from an Excel spreadsheet ("Counties" sheet) and merges it into Word template files by replacing `{FIELD}` placeholders with actual values. Produces both .docx and .pdf output files organized in an Output directory.
- **Max Row Length** - Checks each line of a label template after field substitution to make sure it fits on the label without wrapping. Every line is measured piece by piece in the real font, size, and weight of each run on it, and compared to the label cell's usable width read from the docx itself. Lines that will wrap are flagged with a suggested smaller point size.
- **Check URLs** - Validates Board of Elections (BOE) URLs from the spreadsheet, flagging unreachable or broken links. Optionally opens URLs in a browser for manual review.
- **Check Short URLs** - Compares shortened URLs (Rebrandly, Bitly, etc.) against their expected full URLs to verify they redirect correctly.

Python uv
---------

[uv](https://docs.astral.sh/uv/) is a fast Python package and project manager. It replaces `pip`, `venv`, and other tools with a single command. LeviTool uses uv to manage its dependencies and Python version.

### Install uv on Mac

**Option 1 - Install script** (recommended):

Open Terminal and run:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**Option 2 - Homebrew** (if you already use Homebrew):

```bash
brew install uv
```

After installing, close and reopen your terminal. Verify it worked with:

```bash
uv --version
```

### Update uv

To update uv to the latest version later:

```bash
uv self update
```

Or with Homebrew: `brew upgrade uv`

### Common uv Commands

| Command | What it does |
|---------|-------------|
| `uv sync` | Install all project dependencies (run this after cloning) |
| `uv run python LeviTool.py` | Run a Python script using the project's environment |
| `uv add <package>` | Add a new package to the project |
| `uv pip list` | List all installed packages |
| `uv python list` | Show available Python versions |

### First-Time Setup

Once uv is installed, navigate to the LeviTool project directory and run:

```bash
uv sync
```

This will automatically download the correct Python version and install all required packages into a `.venv` folder inside the project. You don't need to install Python separately - uv handles it for you.

### Set Up uv in PyCharm

PyCharm needs to know which Python interpreter to use. After running `uv sync`, a `.venv` folder is created in the project. Point PyCharm to the Python inside it:

1. Open your project in PyCharm
2. Go to **PyCharm > Settings** (or press `Cmd + ,`)
3. Navigate to **Project: uv_levitool > Python Interpreter**
4. Click the gear icon and choose **Add Interpreter > Add Local Interpreter**
5. Select **Existing** on the left
6. Click the `...` button and browse to the project's `.venv/bin/python` file (e.g., `/path/to/uv_levitool/.venv/bin/python`)
7. Click **OK** to confirm

PyCharm will now use the same Python and packages that uv installed. You should see all the project's dependencies listed in the interpreter panel.

**Tip:** If you ever run `uv sync` again from the terminal (e.g., after pulling new code), PyCharm will automatically pick up the changes since it's pointing at the same `.venv`.

GitHub
------

[GitHub](https://github.com) is a website that hosts code repositories. LeviTool's code is stored on GitHub, which makes it easy to download and stay up to date.

### Install Git on Mac

Git may already be installed. Check by opening Terminal and running:

```bash
git --version
```

If it's not installed, macOS will prompt you to install the Xcode Command Line Tools. Follow the prompts, or install manually:

```bash
xcode-select --install
```

### Download LeviTool

1. Open Terminal
2. Navigate to the folder where you want to put the project, for example:
   ```bash
   cd ~/Documents
   ```
3. Clone the repository:
   ```bash
   git clone https://github.com/kramsman/uv_levitool.git
   ```
4. Install dependencies:
   ```bash
   cd levitool-root
   uv sync
   ```

### Open in PyCharm

1. Open PyCharm
2. Choose **File > Open**
3. Navigate to the `uv_levitool` folder you just cloned and click **Open**
4. PyCharm may prompt you to configure a Python interpreter. Select the one in the `.venv` folder that `uv sync` created (e.g., `.venv/bin/python`)
5. Right-click `LeviTool.py` in the project panel and choose **Run 'LeviTool'**

## Requirements

- Python 3.12+
- macOS (uses `osascript` to quit Microsoft Word, and `docx2pdf` for PDF conversion via Word)
- Microsoft Word (for PDF conversion)

## Usage

A menu will appear with the options: **LeviTool**, **Max Row Len**, **Check Urls**, **Check Short Urls**, and **Exit**.

### Input Structure

Place all input files in a directory named `Input`:
- One `.xlsx` file with a "Counties" sheet (row 1 is ignored, row 2 has `{KEY}` field names, remaining rows have data)
- One or more `.docx` template files containing `{FIELD}` placeholders to be replaced
- A `{USE}` column in the spreadsheet controls which rows are processed
- `{STATE}` and `{CNTYFILENAME}` columns are required for output file naming

### Output

Output is written to an `Output` directory (sibling to `Input`) with subdirectories:
- `Output/Docxs/` - merged Word documents
- `Output/Pdfs/` - PDF conversions of the merged documents
