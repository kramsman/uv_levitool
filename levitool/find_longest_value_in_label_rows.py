""" Check the lengths of row text in a label after Levitool substitutions to make sure it will fit.
Should be max of 34.
"""

# FIXME check if {priority} changes the number of rows used in count

from copy import deepcopy
from pathlib import Path

from docx import Document  # package in Conda is python-docx, not simply docx
from uvbekutils import pyautobek
from uvbekutils import safe_str
from uvbekutils import scroll_box
from uvbekutils import list_pick
from uvbekutils import standardize_columns

from .read_boe_xls import read_boe_xls
from .constants import DEFAULT_FONT_SIZE_PT
from .label_width import resolve_font, width_pt, max_fitting_size_pt, DEFAULT_FAMILY


def get_label_text(label_docx) -> str:
    """Reads text from the upper-left cell of the first table in a label docx file.

    Args:
        label_docx: Path or string path to the label docx template file.

    Returns:
        str: Text from the upper-left label cell with lines joined by newline characters.
    """

    # string or path?
    document_of_docx_file = Document(label_docx)
    d = deepcopy(document_of_docx_file)

    # upper left label
    lines = [line.text for line in d.tables[0].rows[0].cells[0].paragraphs]
    label_text = '\n'.join(lines)
    print(f"{label_text=}")
    return label_text


def _para_size(para):
    """Returns (size_pt, estimated) for a paragraph.

    Reads the Latin font size (w:sz) from the run, then the paragraph style, then the
    paragraph mark. When no Latin size is set, falls back to the complex-script size
    (w:szCs) and finally to DEFAULT_FONT_SIZE_PT; estimated is True in that case
    because the value was inferred, not read from an explicit Latin size.
    """
    from docx.oxml.ns import qn

    size_pt = None
    # 1. Explicit Latin size on a run.
    for run in para.runs:
        if run.font.size is not None:
            size_pt = run.font.size.pt
            break
    # 2. The paragraph's named style.
    if size_pt is None and para.style is not None and para.style.font.size is not None:
        size_pt = para.style.font.size.pt
    # 3. Latin size on the paragraph mark (w:pPr/w:rPr/w:sz, in half-points).
    if size_pt is None:
        sz = para._p.find(qn('w:pPr') + '/' + qn('w:rPr') + '/' + qn('w:sz'))
        if sz is not None and sz.get(qn('w:val')) is not None:
            size_pt = int(sz.get(qn('w:val'))) / 2

    estimated = False
    if size_pt is None:
        # No Latin size anywhere: fall back to the complex-script size (w:szCs),
        # then to the default. These are estimates, not the real Latin size.
        estimated = True
        szcs = next(iter(para._p.iter(qn('w:szCs'))), None)
        if szcs is not None and szcs.get(qn('w:val')) is not None:
            size_pt = int(szcs.get(qn('w:val'))) / 2  # half-points -> points
        else:
            size_pt = DEFAULT_FONT_SIZE_PT

    return size_pt, estimated


def _para_bold(para) -> bool:
    """True when any text-bearing run in the paragraph is bold.

    Word splits a line (and even a single {token}) across several runs with mixed
    formatting, so rather than measure each run separately we treat the whole line as
    bold if any of its visible text is bold. Bold text is wider, so this errs toward
    warning about overflow (the safe direction) and matches the simplification that a
    partly-bold line is measured as fully bold.
    """
    style_bold = bool(para.style and para.style.font.bold)
    for run in para.runs:
        if not run.text.strip():
            continue
        b = run.font.bold
        if b is None:
            b = style_bold
        if b:
            return True
    return False


def _para_font(para):
    """Returns the dominant font family name for a paragraph, or None if unset.

    Weights each run's font by its text length and returns the family covering the
    most characters, so a short differently-fonted run does not swing the result.
    Returns None when no run and no style names a font (the caller then assumes the
    default family).
    """
    counts = {}
    for run in para.runs:
        n = len(run.text)
        if n == 0:
            continue
        name = run.font.name
        if name is None and para.style is not None:
            name = para.style.font.name
        if name:
            counts[name] = counts.get(name, 0) + n
    if counts:
        return max(counts, key=counts.get)
    if para.style is not None and para.style.font.name:
        return para.style.font.name
    return None


def get_label_line_font_info(label_docx) -> list:
    """Returns per-line font info for the label cell used to measure real line width.

    Args:
        label_docx: Path or string path to the label docx template file.

    Returns:
        list: One dict per paragraph in the label cell, in order, with keys:
            'line' (1-based number), 'size' (point size, float), 'estimated' (True when
            the size was inferred rather than read from an explicit Latin size), 'bold'
            (True when any visible text on the line is bold), and 'font' (dominant font
            family name, or None when none is specified).
    """
    d = deepcopy(Document(label_docx))
    info = []
    for i, para in enumerate(d.tables[0].rows[0].cells[0].paragraphs):
        size_pt, estimated = _para_size(para)
        info.append({
            'line': i + 1,
            'size': size_pt,
            'estimated': estimated,
            'bold': _para_bold(para),
            'font': _para_font(para),
        })
    return info


def get_label_cell_usable_width_pt(label_docx) -> float:
    """Returns the usable text width (in points) of the label's upper-left cell.

    Usable width is the cell width minus its left and right margins. The cell width is
    read from w:tcW (falling back to the table grid column), and margins from the cell's
    own w:tcMar, then the table default w:tblCellMar; when no margin is specified,
    Word's default of 108 twips (0.075 in) per side is assumed. Returns None if the cell
    width cannot be determined. Twips are converted at 20 twips per point.

    Args:
        label_docx: Path or string path to the label docx template file.

    Returns:
        float: Usable cell width in points, or None if it cannot be determined.
    """
    from docx.oxml.ns import qn

    # Word's default table cell side margin when none is specified.
    DEFAULT_SIDE_MARGIN_TWIPS = 108

    d = deepcopy(Document(label_docx))
    tbl = d.tables[0]
    cell = tbl.rows[0].cells[0]
    tcPr = cell._tc.find(qn('w:tcPr'))

    def _w(element):
        """Reads a w:w attribute (twips) off an element, or None."""
        if element is None:
            return None
        val = element.get(qn('w:w'))
        return int(val) if val is not None else None

    # Cell width: prefer the cell's own w:tcW, else the first grid column.
    cell_w = _w(tcPr.find(qn('w:tcW'))) if tcPr is not None else None
    if cell_w is None:
        grid_col = tbl._tbl.find(qn('w:tblGrid') + '/' + qn('w:gridCol'))
        cell_w = _w(grid_col)
    if cell_w is None:
        return None

    def _side_margin(container, mar_tag, side):
        """Reads a left/right margin (twips) from a *Mar element, or None."""
        if container is None:
            return None
        mar = container.find(qn(mar_tag))
        if mar is None:
            return None
        return _w(mar.find(qn('w:' + side)))

    tblPr = tbl._tbl.tblPr
    left = _side_margin(tcPr, 'w:tcMar', 'left')
    if left is None:
        left = _side_margin(tblPr, 'w:tblCellMar', 'left')
    if left is None:
        left = DEFAULT_SIDE_MARGIN_TWIPS
    right = _side_margin(tcPr, 'w:tcMar', 'right')
    if right is None:
        right = _side_margin(tblPr, 'w:tblCellMar', 'right')
    if right is None:
        right = DEFAULT_SIDE_MARGIN_TWIPS

    usable_twips = cell_w - left - right
    return usable_twips / 20.0  # 20 twips per point


def max_label_lengths(*, used_field: str = None, label_docx=None, initial_attachment_dir=None) -> None:
    """Checks the maximum line length in a label template after mail-merge substitutions.

    Prompts for the label docx and auto-detects the BOE xlsx from the same directory.
    Substitutes all {KEY} tokens for every county row, then reports the widest line
    per label line number, measuring each line's real rendered width (actual font,
    size, and weight) against the label cell's usable width. Displays results and any
    overflow warnings in a scroll box.

    Args:
        used_field (str): Column key identifying active county rows (e.g. '{priority}').
            Prompts if None.
        label_docx: Path to the label docx template. Prompts via file picker if None.
        initial_attachment_dir: Starting directory for the label docx file picker.
            Used only when label_docx is None.
    """

    import pandas as pd
    from pathlib import Path
    from loguru import logger
    import re
    import glob
    from uvbekutils import exit_yes, select_file

    def print_max_line_info(df: pd.DataFrame, line_list_field: str, line_fonts: list = None,
                            usable_width_pt: float = None) -> str:
        """Finds and prints the widest substituted line for each line position in the label.

        Measures each substituted line's real rendered width in its actual font family,
        point size, and weight, then compares against the label cell's usable width and
        appends a warning (with a suggested smaller size) if the line will not fit. The
        widest line is chosen by measured width, not character count, because fonts are
        proportional.

        Args:
            df (pd.DataFrame): DataFrame where each row contains a list of substituted
                label lines in the column specified by line_list_field.
            line_list_field (str): Name of the column containing per-row lists of label lines.
            line_fonts (list): Optional list of per-line font-info dicts from
                get_label_line_font_info(), giving each line's size, bold, and font family.
            usable_width_pt (float): Usable cell width in points from
                get_label_cell_usable_width_pt(); the fit limit for every line.

        Returns:
            str: Newline-joined summary of the widest line, its size/font/width, and any
                overflow warning for each label line position, or 'None' if empty.
        """
        if df.empty:
            pyautobek.alert(f"No counties are being selected based on the field '{used_field}'.\n\n"
                            f"Checking for these copunties will be skipped.\n"
                            f"If desired, rerun choosing another field (like {{priority}}'" ,
                        "No Counties Being Selected")
            result = 'None'
        else:
            # get number of list elements (# of lines) in the label, so we know how many need to be checked
            number_of_lines = len(df[line_list_field].iloc[0])
            result_lines = []
            for line_number in range(number_of_lines):
                line_data = [line[line_number] for line in df[line_list_field]]

                info = line_fonts[line_number] if line_fonts and line_number < len(line_fonts) else {}
                size_pt = info.get('size')
                estimated = info.get('estimated', False)
                bold = info.get('bold', False)
                font = info.get('font')
                font_assumed = font is None  # no font named in the docx; default assumed

                # Resolve the real font file once; measure every candidate line's width
                # and keep the widest by RENDERED width (not char count, since fonts are
                # proportional). font_path is None only if no usable font was found.
                font_path, family_used, fallback = resolve_font(font, bold)
                measured = None
                if size_pt and font_path:
                    widths = [(ln, width_pt(font_path, ln, size_pt)) for ln in line_data]
                    max_line, measured = max(widths, key=lambda t: t[1])
                else:
                    max_line = max(line_data, key=len)

                size_num = f"{size_pt:g}" if size_pt else ""
                weight_str = " bold" if bold else ""
                # Note when the family or size had to be assumed rather than read.
                notes = []
                if estimated:
                    notes.append("size est.")
                if font_assumed or fallback:
                    notes.append(f"font assumed {family_used}")
                note_str = f" ({', '.join(notes)})" if notes else ""
                font_str = f" {family_used}" if size_pt else ""

                if measured is not None and usable_width_pt:
                    over = measured > usable_width_pt
                    meta_str = (f"  - {size_num} pt{weight_str}{font_str}, "
                                f"{measured:.0f} of {usable_width_pt:.0f} pt wide{note_str}")
                    if over:
                        suggested = max_fitting_size_pt(font_path, max_line, size_pt, usable_width_pt)
                        if suggested and suggested >= 5:
                            fit_str = f"try {suggested:g} pt"
                        else:
                            fit_str = "shorten text - won't fit"
                        warning = f"   - **** TOO WIDE - {fit_str} ****"
                    else:
                        warning = "   - fits"
                else:
                    # No font metrics available (font size unknown or no font found):
                    # report the longest line by character count without a verdict.
                    meta_str = f"  - {len(max_line)} chars (width not measured{note_str})"
                    warning = ""

                line_info = f"Line {line_number + 1},  '{max_line}'{meta_str}{warning}"
                print(line_info)
                result_lines.append(line_info)
                result = '\n'.join(result_lines)

        return result

    if label_docx is None:
        label_docx = select_file("PICK LABEL DOCX",
                                 start_dir=initial_attachment_dir,
                                 files_like='*.docx',
                                 mode='file',
                                 title2="Pick label docx template that will have BOE info merged into it "
                                        "('LABEL 30 per page...')",
                                 )
    logger.info(f"Label docx being used for input: '{label_docx}")
    label_docx = Path(label_docx).expanduser()

    label_text = get_label_text(label_docx)
    label_line_fonts = get_label_line_font_info(label_docx)
    usable_width_pt = get_label_cell_usable_width_pt(label_docx)
    logger.info(f"Label cell usable width: {usable_width_pt} pt")
    keys = find_keys_in_text(label_text)
    keys = [key.lower() for key in keys]

    # list of all xlsx files not starting with ~ (temporary files)
    xlsx_files = glob.glob(str(label_docx.parent / '[!~]*.xlsx'))

    if len(xlsx_files) != 1:  # must have 1, not more
        msg = f"Can not select BOE xlsx file from label directory because there are more than one;" \
              f"there are {len(xlsx_files)}.  EXITING."
        exit_yes(msg, "MORE THAN ONE XLSX")
    # xlsx_file_name = xlsx_files[0]  # filename of the first (and only) xlsx
    boe_xls = Path(xlsx_files[0]).expanduser()

    county_sheet_keys, df = read_boe_xls(boe_xls)

    # Lowercase all column names to match lowercase keys
    df.columns = [col.lower() for col in df.columns]

    if used_field is None:
        used_field = list_pick(lst=df.columns,
                               title="PICK KEY FIELD TO ID 'USED' COUNTIES",
                               msg="All counties with this field not empty will be checked as a group ("
                                   "'{priority}' will be used if none chosen).",
                               select_mode='single',
                               pre_select=False,
                               allow_none=True,
                               )[0]
        if used_field == '':
            used_field = '{priority}'

    logger.info(f"Label docx being used for input: '{label_docx}")


    # set object type so field can accept an object - a list
    df['lines'] = ''
    df['lines'] = df['lines'].astype('object')

    for index, row in df.iterrows():
        tmp_label_text = label_text  # so we start fresh with the label info containing the {} keys

        for fld in keys:
            # replace occurrences of the key with the value in the df (ex '{county}' gets replaced with 'Dade')
            # tmp_label_text = tmp_label_text.replace(fld, df.at[index, fld])
            # TODO what ot do if nan?  takes up only one space but might be much larger when filled in later.
            tmp_label_text = re.sub(fld, safe_str(df.at[index, fld]), tmp_label_text, flags=re.IGNORECASE)

        line_list = safe_str(tmp_label_text).split('\n')  # create list of lines from the text
        line_list = [line.strip('\t ') for line in line_list if len(line.strip('\t ')) > 0]  # trim and remove blanks

        # fill a field in the df with the list of line text
        df.at[index, 'lines'] = line_list

    # Collect max line info for alert display
    max_line_results = []
    total_rows = len(df)

    # Below is True if 'nan' found in any cell in 'keys' columns
    nan_found_in_all_rows = (df[keys] == 'nan').any(axis="columns").any(axis="rows")
    # Below creates a boolean series/mask: (df[keys] == 'nan').any(axis="columns")
    # df[keys] selects only columns in list 'keys'
    # below puts it together: creates df with 'keys' columns and rows containing 'nan' in any column:
    #   df[keys][(df[keys] == 'nan').any( axis="columns")]
    df_nan = df[keys][(df[keys] == 'nan').any(axis="columns")]

    if nan_found_in_all_rows:
        print("\n20 rows containing 'nan' somewhere in a ALL row 'key' columns")
        with pd.option_context('display.max_rows', 20, 'display.max_columns', None, ):
            print(df_nan)
        pyautobek.alert(f"'Nan' (empty) found in {len(df_nan)} rows in at least one key of '{keys}' to be substituted "
                        f"in ALL rows so analysis of all rows skipped.\n\nSee log for print of 20 rows.",
                        "'nan' found somewhere in ALL row keys")

        # there are some keys with nan so can't run on all.  Instead, find those rows with all keys non-nan.
        df_non_nan = df[(df[keys].notnull()).all(axis="columns")]  # note all columns taken
        if len(df_non_nan) != 0:
            print("ALL ROWS WITH NO NAN VALUES")
            print(df_non_nan[keys])
            print()
            print("MAX LINES IN SUBSTITUTED SCRIPT INFO FOR ALL ROWS WITH NO NAN VALUES")
            result = print_max_line_info(df_non_nan, 'lines', line_fonts=label_line_fonts,
                                         usable_width_pt=usable_width_pt)
            max_line_results.append(("ALL ROWS WITH NO NAN VALUES", result, len(df_non_nan)))
        print()
        a = 1
    else:
        print("MAX LINES IN SUBSTITUTED SCRIPT INFO FOR ALL ROWS")
        result = print_max_line_info(df, 'lines', line_fonts=label_line_fonts,
                                     usable_width_pt=usable_width_pt)
        max_line_results.append(("ALL ROWS", result, len(df)))
        print()

    # A row is "used" when its used_field is genuinely non-blank: not empty,
    # not whitespace, and not a missing value (read as <NA> or the literal 'nan').
    used_vals = df[used_field].fillna('').astype(str).str.strip()
    used_mask = (used_vals != '') & (used_vals.str.lower() != 'nan')

    # see explanation of expression above
    # Check for bad row where NAN is found somewhere in a key field.
    nan_found_in_select_rows = (df[keys].loc[used_mask] == 'nan') \
        .any(axis="columns").any(axis="rows")
    df_nan = df[keys][used_mask & (df[keys] == 'nan').any(axis="columns")]
    if nan_found_in_select_rows:
        print(f"\nThe following rows with non-blank '{used_field}' contain 'nan' somewhere in ALL 'key' columns")
        with pd.option_context('display.max_rows', None, 'display.max_columns', None, ):
            print(df_nan)

        print(f"\nSTATS NOT PRODUCED - USED ROWS HAVE BAD 'nan' DATA")

        pyautobek.alert(
            f"'Nan' (blank) found in {len(df_nan)} rows at least one key with non-blank '{used_field}' to be substituted so analysis "
            f"of all rows skipped.\n\n"
            "See log for list for rows.",
            "'nan' found in USED row keys")
    else:
        print(f"MAX LINES IN SUBSTITUTED SCRIPT INFO FOR USED SUB ROWS ('{used_field}' not blank)")
        df_used_counties = df.loc[used_mask]
        result = print_max_line_info(df_used_counties, 'lines', line_fonts=label_line_fonts,
                                     usable_width_pt=usable_width_pt)
        max_line_results.append((f"'USED' COUNTIES ('{used_field}' not blank)", result, len(df_used_counties)))
        print()

    print("RAW LABEL DATA:")
    print(label_text)

    # Build and display alert with max line lengths and label text
    alert_lines = []
    alert_lines.append("--- LABEL TEMPLATE TEXT ---")
    alert_lines.append(label_text)
    alert_lines.append("")
    for section_name, section_result, row_count in max_line_results:
        alert_lines.append(f"--- {section_name} ---")
        alert_lines.append(section_result)
        alert_lines.append(f"({row_count} rows of {total_rows})")
        alert_lines.append("")

    alert_lines.append(f"'Used' counties based on field {used_field}")
    alert_lines.append("")
    alert_lines.append(f"Label File: {label_docx}")
    alert_lines.append("")
    alert_lines.append(f"BOE Sheet: {boe_xls}")
    alert_lines.append("")

    if usable_width_pt:
        alert_lines.append(
            f"Fit check: each line's real rendered width (its actual font, size, and\n"
            f"bold) is measured and compared to the label cell's usable width of\n"
            f"{usable_width_pt:.1f} pt ({usable_width_pt / 72:.2f} in). A line marked TOO WIDE will\n"
            f"wrap in Word; try the suggested smaller size or shorten the text.")
    else:
        alert_lines.append(
            "Fit check: the label cell width could not be read, so line widths were\n"
            "not measured. Character counts are shown for reference only.")

    alert_message = '\n'.join(alert_lines)
    # pyautobek.alert(alert_message, "MAX LABEL LINE LENGTHS")
    scroll_box(alert_message, title="MAX LABEL LINE LENGTHS", wrap_lines=False, width=1200, height=600)


def find_keys_in_text(label_text: str) -> set:
    """Finds all {KEY} replacement tokens in a plain text string.

    Args:
        label_text (str): Text content from a label template, typically from get_label_text().

    Returns:
        set: Set of key strings found (e.g. {'{county}', '{phone}'}).
    """

    import re
    keys = set(re.findall(r'{[a-zA-Z0-9]+}', safe_str(label_text)))
    return keys


if __name__ == '__main__':
    INITIAL_ATTACHMENT_DIR = Path("~/Dropbox/Postcard Files/Attachments/Campaigns").expanduser()
    # LENGTH_CHECK_FIELDS_SELECT = '{priority}'
    LENGTH_CHECK_FIELDS_SELECT = '{use}'

    # Interactive GUI call (file picker + field picker dialogs). Use Run (not Debug):
    # Qt modal dialogs freeze under PyCharm's debugger on Python 3.12.
    max_label_lengths(used_field=None, initial_attachment_dir=INITIAL_ATTACHMENT_DIR, )

    # DEBUG (dialog-free) alternative — hard-code inputs, skips select_file / list_pick:
    # max_label_lengths(
    #     used_field='{use}',
    #     label_docx=Path("~/Dropbox/Postcard Files/Attachments/Campaigns/GA General 6-2026/Input/"
    #                     "LABELS-30 per page GA General 6-15-26.docx").expanduser(),
    # )

    a = 1
    # Label text is copied here.  Lines are trimmed.
    #     LABEL_TEXT = """
    # {cntytoprint} Registrar
    # Phone Num: {specialphone}
    # or {specialurl}
    # Early Voting Now thru Feb 18
    #     """

    # 30 per page, Avery 5161 Label
    # Font Size | Approximate Characters per Line
    # ----------|--------------------------------
    # 8 pt      | 52-55 characters
    # 9 pt      | 46-49 characters
    # 10 pt     | 42-45 characters
    # 11 pt     | 38-41 characters
    # 12 pt     | 35-38 characters
    # 13 pt     | 32-35 characters
    # 14 pt     | 30-33 characters

    # 20 per page, Avery 5161 Label
    # Font Size | Approximate Characters per Line
    # ----------|--------------------------------
    # 8 pt      | 80-85 characters
    # 9 pt      | 71-76 characters
    # 10 pt     | 64-68 characters
    # 11 pt     | 58-62 characters
    # 12 pt     | 53-57 characters
    # 13 pt     | 49-52 characters
    # 14 pt     | 46-49 characters

    # label_text_file = Path("~/Dropbox/Postcard Files/Attachments/Campaigns/TEST NAN3/Input/label.txt").expanduser()
    # with open(label_text_file, "r") as f:
    #     LABEL_TEXT = f.read()

    # LENGTH_CHECK_FIELDS_SELECT = '{priority}'
    # LENGTH_CHECK_FIELDS = [['county', 'phone'], 'county', 'phone']
    # LENGTH_CHECK_FIELDS = ['{cntytoprint}', '{specialphone}', '{specialurl}']
    # test = ['+'.join(x) for x in LENGTH_CHECK_FIELDS]
    # test = ['county', 'phone']

    # data = [['Small', '555-555-5555', 'x'], ['Medium', '555-555-5552 x2', 'x'], ['VeryVeryLonnnng', '555-555-5553 ext 3', '']]
    # df = pd.DataFrame(data, columns=['{county}', '{phone}', '{PRIORITY}'])

    # max_label_lengths(initial_attachment_dir=INITIAL_ATTACHMENT_DIR,)
    # label_docx="~/Dropbox/Postcard Files/Attachments/Campaigns/TEST NAN3/Input/TEST Special GOTV "
    #            "LABELS-30per page 1-3-2023.docx")
    # a = 1

    # boe_xls = "~/Dropbox/Postcard Files/Attachments/Campaigns/TEST NAN3/Input/BOE Info VA (1).xlsx",

    # df_to_check_length = pd.DataFrame
    #
    # fieldnames = []
    # for item in LENGTH_CHECK_FIELDS:
    #     if isinstance(item, list):
    #         fieldname = '+'.join(item)
    #         df_to_check_length[fieldname] = df[item].astype(str).apply(lambda x: ''.join(x), axis=1)
    #     elif isinstance(item, str):
    #         df_to_check_length[item] = df[item]
    #     else:
    #         print('ERROR - field not on dataframe')
    #     fieldnames.append(fieldname)
    #
    # df_to_check_length[LENGTH_CHECK_FIELDS_SELECT] = df_to_check_length[LENGTH_CHECK_FIELDS_SELECT]
    #
    #
    # for field in LENGTH_CHECK_FIELDS:
    #     # max_field = max(df[f"{{{field}}}"], key=len)
    #     max_field = max(df[f"{field}"], key=len)
    #     print(f"Max of '{field}' column: '{max_field}', length is {len(max_field)}")
    #
    # a=1

    # def fill_lines(df, field_list, used_field):
    #
    #     # set object type so field can accept an object - a list
    #     df['lines'] = ''
    #     df['lines'] = df['lines'].astype('object')
    #
    #     for fld in field_list:
    #         locals()[fld] = fld
    #         # newfield = df[fld]
    #
    #     for index, row in df.iterrows():
    #         # county = row['county']
    #         # phone = row['phone']
    #         # lines.append(f"{county} is the county name.")
    #
    #         df.at[index, 'lines'] = \
    #             [
    #                 f"{county} is the county name.",
    #                 f"{phone} is the phone.",
    #                 f"Together we have {county} and {phone}.",
    #                 f"Again we have {county} and {phone}."
    #             ]
    #
    #     entries = len(df.at[0, 'lines'])
    #
    #     for entry in range(entries):
    #         line_data = [x[entry] for x in df['lines'] ]
    #         max_line = max(line_data, key=len)
    #         max_length = len(max_line)
    #         print(f"line_data: {line_data}")
    #         print(f"max_line: '{max_line}', len:{max_length}\n")
    #
    #
    #     # for entry in range(entries):
    #     #     # max_field = max(df[f"{{{field}}}"], key=len)
    #     #     max_field = max(df['line'][entry], key=len)
    #     #     print(f"Max of '{field}' column: '{max_field}', length is {len(max_field)}")
    #
    a = 1
