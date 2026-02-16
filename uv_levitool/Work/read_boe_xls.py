def read_boe_xls(boe_xls_file, key_check_list=None, display_keys=False):
    """ reads the key fields and data from boe_xls files used by the LeviTool
    returning the keys in a list and the data in a dataframe.
    All keys converted to lowercase.
    Can check for occurrences of keys, usually erroring if more than one.
    Columns with blank keys are removed.  Keys not surrounded by {} are errored.  First row is ignored."""

    def check_count(key_str, key_count, oper, num_allowed):
        """ make sure headers appear in file only given number of times and error if not

        Args:
            key_str (): key being checked, eg '{use}'
            key_count (): number of key_str found
            oper (): comparison passed as function from operator module eg operator.gt.  https://stackoverflow.com/questions/18591778/how-to-pass-an-operator-to-a-python-function
            num_allowed (): comparison value
            :param key_str:
            :type key_str:
            :param key_count:
            :type key_count:
            :param oper:
            :type oper:
            :param num_allowed:
            :type num_allowed:
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
