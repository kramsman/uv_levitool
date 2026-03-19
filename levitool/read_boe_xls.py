def read_boe_xls(boe_xls_file, key_check_list=None, display_keys: bool = False) -> tuple:
    """Reads key fields and county data from a BOE xlsx file used by LeviTool.

    Reads the 'Counties' sheet, extracts column header keys from row 2 (row 1 is
    skipped), validates that all keys are enclosed in curly braces and appear exactly
    once, removes columns with blank keys, and returns the keys list and the full data
    DataFrame with values stripped of whitespace.

    Args:
        boe_xls_file: Path or string path to the BOE xlsx file.
        key_check_list: Unused; reserved for future key validation. Defaults to None.
        display_keys (bool): If True, display the keys list. Defaults to False.

    Returns:
        tuple: A (keys, county_sheet) pair where keys is a list of lowercase key strings
            (e.g. ['{use}', '{state}', '{county}']) and county_sheet is a DataFrame
            with all values read as strings and whitespace stripped.
    """

    def check_count(key_str: str, key_count: int, oper, num_allowed: int) -> None:
        """Validates that a key count satisfies a given comparison; exits on failure.

        Args:
            key_str (str): The key being checked (e.g. '{use}').
            key_count (int): The actual number of occurrences found.
            oper: A comparison function from the operator module (e.g. operator.eq).
                See https://stackoverflow.com/questions/18591778
            num_allowed (int): The value to compare against.
        """

        if not oper(key_count, num_allowed):
            msg1 = f"Key '{key_str}' count of {key_count} is not {oper} {num_allowed}."
            exit_yes(msg1, "ERROR")

    import pandas as pd
    from loguru import logger
    # from collections import Counter
    import operator

    from uvbekutils import exit_yes, exit_yes_no

    county_sheet_name = 'Counties'
    key_row = 1  # zero indexed

    # get keys: read only one row after skipping if needed
    # Note .astype(str) below which reads all data as string and converts missing to 'nan'
    county_sheet_temp = pd.read_excel(boe_xls_file, sheet_name=county_sheet_name, header=None, skiprows=key_row,
                                      nrows=1).astype(str)
    # get keys dataframe row
    keys = county_sheet_temp.loc[0, :].values.tolist()
    key_count_before = len(keys)
    import pandas as pd

    keys = [x.strip().lower() for x in keys if not pd.isna(x)]
    if len(keys) != key_count_before:
        logger.warning(f"{key_count_before - len(keys)} columns had blanks in key row.  Is tht ok?")

    # check key begins and ends with {}
    for key in keys:
        if not (key.startswith("{") and key.endswith("}")):
            msg1 = f"Key '{key}' does not start with '{{' or end with '}}'."
            exit_yes(msg1, "ERROR")

    # check all keys appear only once
    for key in keys:
        check_count(key, keys.count(key), operator.eq, 1)

    # display keys if prompted

    # read data in county sheet
    try:
        county_sheet = pd.read_excel(boe_xls_file, sheet_name=county_sheet_name, header=0, skiprows=key_row).astype(str)
    except BaseException as e:
        msg = "Install openpyxl 3.0.10 not 3.1.2 because later bombs if xlsx has filters"
        logger.error(msg)
        logger.error(e)
        logger.error("Install openpyxl 3.0.10 not 3.1.2 because later bombs if xlsx has filters")
        exit_yes(msg)

    # convert all column names to lowercase
    county_sheet.columns = [col.lower() for col in county_sheet.columns]

    # remove columns which had blank keys (appear as 'unnamed: [colnumber])
    for column in county_sheet.columns:
        if 'unnamed:' in column:
            del county_sheet[column]

    # replace fields with stripped values then "' '" to account for and see multi spaces
    county_sheet = county_sheet.apply(lambda x: x.str.strip())

    return keys, county_sheet


if __name__ == '__main__':
    from uvbekutils import setup_loguru
    from pathlib import Path

    setup_loguru("DEBUG", "DEBUG")

    k, c = read_boe_xls(Path("~/Dropbox/Postcard Files/PythonPrograms/Development/LeviTool/Work/TEST BOE Info "
                             "VA.xlsx").expanduser(),
                        )

    a = 1
