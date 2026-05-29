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
from .constants import MAX_CHARS_PER_LINE_TEXT, CHARS_PER_LINE_30PP


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


def get_label_font_sizes(label_docx) -> list:
    """Returns the detected font size (in points) for each paragraph line in the label cell.

    Checks each run's font size first, then falls back to the paragraph style font size.
    Returns None for a line if the size cannot be determined (inherited from document default).

    Args:
        label_docx: Path or string path to the label docx template file.

    Returns:
        list: List of (line_number, font_size_pt) tuples, one per paragraph in the label cell.
            font_size_pt is a float or None if undetermined.
    """
    d = deepcopy(Document(label_docx))
    default_size_pt = None
    try:
        default_size = d.styles['Normal'].font.size
        if default_size is not None:
            default_size_pt = default_size.pt
    except Exception:
        pass

    font_sizes = []
    for i, para in enumerate(d.tables[0].rows[0].cells[0].paragraphs):
        size_pt = None
        for run in para.runs:
            if run.font.size is not None:
                size_pt = run.font.size.pt
                break
        if size_pt is None and para.style.font.size is not None:
            size_pt = para.style.font.size.pt
        font_sizes.append((i + 1, size_pt))

    # Fill None entries using any detected size from other lines
    detected = next((sz for _, sz in font_sizes if sz is not None), default_size_pt)
    font_sizes = [(n, sz if sz is not None else detected) for n, sz in font_sizes]
    return font_sizes


def max_label_lengths(*, used_field: str = None, label_docx=None, initial_attachment_dir=None) -> None:
    """Checks the maximum line length in a label template after mail-merge substitutions.

    Prompts for the label docx and auto-detects the BOE xlsx from the same directory.
    Substitutes all {KEY} tokens for every county row, then reports the longest line
    per label line number. Displays results and character-limit reference tables for
    common Avery label formats in a scroll box.

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

    def print_max_line_info(df: pd.DataFrame, line_list_field: str, font_sizes: list = None) -> str:
        """Finds and prints the longest substituted line for each line position in the label.

        Compares max line length against the Avery 5160 30per page character limit for the
        detected font size and appends a warning if the line may run over.

        Args:
            df (pd.DataFrame): DataFrame where each row contains a list of substituted
                label lines in the column specified by line_list_field.
            line_list_field (str): Name of the column containing per-row lists of label lines.
            font_sizes (list): Optional list of (line_number, font_size_pt) from
                get_label_font_sizes(). Used to check character limits per line.

        Returns:
            str: Newline-joined summary of the longest line, its length, font size, and
                any overrun warning for each label line position, or 'None' if the DataFrame
                is empty.
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
                max_line = max(line_data, key=len)
                max_line_length = len(max_line)

                size_pt = font_sizes[line_number][1] if font_sizes and line_number < len(font_sizes) else None
                limit = CHARS_PER_LINE_30PP.get(int(size_pt)) if size_pt else None
                warning = "  ** may run over-check" if limit and max_line_length > limit else ""
                if size_pt and limit:
                    size_str = f"  {size_pt} pt ({limit} allowed)"
                elif size_pt:
                    size_str = f"  {size_pt} pt"
                else:
                    size_str = ""

                line_info = f"line {line_number + 1},  max: {max_line_length}   '{max_line}'{size_str}{warning}"
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
    label_font_sizes = get_label_font_sizes(label_docx)
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
            result = print_max_line_info(df_non_nan, 'lines', font_sizes=label_font_sizes)
            max_line_results.append(("ALL ROWS WITH NO NAN VALUES", result, len(df_non_nan)))
        print()
        a = 1
    else:
        print("MAX LINES IN SUBSTITUTED SCRIPT INFO FOR ALL ROWS")
        result = print_max_line_info(df, 'lines', font_sizes=label_font_sizes)
        max_line_results.append(("ALL ROWS", result, len(df)))
        print()

    # see explanation of expression above
    # Check for bad row where NAN is found somewhere in a key field.
    nan_found_in_select_rows = (df[keys].loc[df[used_field].str.strip() != ''] == 'nan') \
        .any(axis="columns").any(axis="rows")
    df_nan = df[keys][(df[used_field].str.strip() != '') & (df[keys] == 'nan').any(axis="columns")]
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
        df_used_counties = df.loc[df[used_field].str.strip() != '']
        result = print_max_line_info(df_used_counties, 'lines', font_sizes=label_font_sizes)
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

    alert_lines.append(MAX_CHARS_PER_LINE_TEXT)

    alert_message = '\n'.join(alert_lines)
    # pyautobek.alert(alert_message, "MAX LABEL LINE LENGTHS")
    scroll_box(alert_message, title="MAX LABEL LINE LENGTHS", wrap_lines=True )


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
    max_label_lengths(used_field=None, initial_attachment_dir=INITIAL_ATTACHMENT_DIR, )

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
