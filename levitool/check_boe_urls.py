
from pathlib import Path
from loguru import logger
from uvbekutils import select_file
from uvbekutils import setup_loguru, exit_yes, list_pick
from uvbekutils import pyautobek

import webbrowser

ATTACHMENTS_INITIAL_DIR = Path("~/Dropbox/Postcard Files/Attachments/Campaigns").expanduser()


def final_url(url: str):
    """Follows all redirects for a URL and returns the final destination URL.

    Useful for comparing shortened URLs to their expected targets.

    Args:
        url (str): The URL to resolve, including any redirects.

    Returns:
        str: The final redirected URL if the request succeeded.
        Exception: The exception raised if the request failed.
    """

    import requests

    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

    try:
        response = requests.get(url, timeout=5, headers=headers)
        return response.url  # final_url
    except Exception as e:
        # logger.debug(e)
        return e


def url_error(url: str):
    """Checks whether a URL is reachable and returns an error description if not.

    Also flags Rebrandly 404 pages (missing short URLs created by Andre).

    Args:
        url (str): The URL to check.

    Returns:
        None: If the URL is reachable (HTTP 200 or 403).
        str: An error description if the URL is unreachable or returns a non-200 status.
        Exception: The requests exception if the connection or timeout failed.
    """

    import requests

    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

    try:
        response = requests.get(url, timeout=5, headers=headers)

        if response.url == 'https://rebrandly.com/404':
            return 'https://rebrandly.com/404'
        elif response.status_code == 200:
            return None
        elif response.status_code == 403:
            return None
        else:
            return f"Response is {response.status_code}, not 200"
    except requests.ConnectionError as e:
        # logger.debug(e)
        return e
    except requests.Timeout as e:
        # logger.debug(e)
        return e


def get_browser_option(open_browser) -> str:
    """Prompts the user to choose a browser-opening mode if one has not already been set.

    Accepted modes are 'no', 'error', and 'all'. If open_browser is True or not one
    of those values, a dialog is shown. Also prompts the user to open a fresh browser
    window if a non-'no' mode is chosen.

    Args:
        open_browser: The current browser option. Pass True or an unrecognized string
            to trigger the prompt dialog; pass 'no', 'error', or 'all' to use as-is.

    Returns:
        str: Lowercase browser option: 'no', 'error', or 'all'.
    """

    from uvbekutils import exit_yes

    if open_browser is True or open_browser.lower() not in ['no', 'error', 'all']:
        # open browser for urls to manually review each one?
        # open_browser = text_box("Do you want to open a browser window for urls?",
        #                         box_title='Open browser?',
        #                         buttons=['No', 'Error', 'All', 'Exit'])
        open_browser = pyautobek.confirm("Do you want to open a browser window for urls?",
                                'Open browser?',
                                buttons=['No', 'Error', 'All', 'Exit'])
        if open_browser == 'exit':
            exit_yes("exit picked")
        # best to open a fresh window because so many tabs are tough to clean up individually
        if open_browser != 'no':
            # text_box("Best to open a fresh browser window now.  Continue?",
            #          box_title='Open Window',
            #          buttons=['Yes'])
            pyautobek.confirm("Best to open a fresh browser window now.  Continue?",
                                    'Open Window',
                                    buttons=['Yes'])
    return open_browser.lower()


def check_boe_urls(key_containing_url: str = None, key_for_used: str = None,
                   open_browser=False, boe_xls=None) -> None:
    """Reads URLs from a BOE xlsx file and reports unreachable ones; optionally opens them in a browser.

    Prompts for the xlsx file, URL key, and 'used' filter key if not provided.
    Checks each unique URL for reachability and optionally opens browser tabs for
    review or error investigation.

    Args:
        key_containing_url (str): Column key holding the URL to check (e.g. '{evurl}').
            Prompts if None.
        key_for_used (str): Column key identifying active county rows (e.g. '{priority}').
            Prompts if None; pass 'all' to check all rows without prompting.
        open_browser: Browser opening mode. True or None prompts; 'all' opens every URL;
            'error' opens only errored URLs; False skips browser entirely.
        boe_xls: Path to the BOE xlsx file. Prompts via file picker if None.
    """

    # FIXME Open browser for all did not open any

    from .read_boe_xls import read_boe_xls

    if boe_xls is None:
        boe_xls = select_file(
            title="Pick BOE file (xlsx) with urls to check.",
            start_dir=ATTACHMENTS_INITIAL_DIR,
            files_like="*.xlsx",
            choices=["Select", "Cancel"],
            mode="file",  # file, dir or both
            title2="Will be in an INPUT subdirectory"
        )
        print(f"Selected: {boe_xls}")

    logger.info(f"BOE xls being checked is: '{boe_xls}")

    county_sheet_keys, county_sheet_df = read_boe_xls(boe_xls)

    # get key like '{url}' if not passed
    if key_containing_url is None:
        # key_containing_url = select_from_list(county_sheet_keys, 'Select tag for url field', select_type='radio')
        key_containing_url = list_pick(county_sheet_keys,
                                       'Select tag for url field',
                                       "Select tag for url field, common choice is '{url}'",
                                       select_mode='single',
                                       pre_select=False)[0]
    # make sure it's on file
    key_containing_url = key_containing_url.lower()
    logger.info(f"Key field being checked is '{key_containing_url}'")
    # make sure key is on file
    if key_containing_url not in county_sheet_keys:
        exit_yes(f"Url key '{key_containing_url}' not in county keys: \n\n{', '.join(county_sheet_keys)}")

    # get key like '{priority}' if not passed
    if key_for_used is None:
        key_for_used = list_pick(county_sheet_keys,
                                 'Select Field',
                                 "Select tag to id counties being used (often '{priority}', None for all)",
                                 select_mode='single',
                                 pre_select=False)[0]

    # make sure key is on file
    if key_for_used is not None:  # None ok from select = all so have to check
        key_for_used = key_for_used.lower()
        if key_for_used not in county_sheet_keys:
            exit_yes(f"Key 'for used' '{key_for_used}' not in county keys: \n\n{', '.join(county_sheet_keys)}")
    logger.info(f"Key field being used to limit counties is '{key_for_used}'")

    open_browser = get_browser_option(open_browser)

    # if a field is specified to identify used rows (those we are interested in) then filter the df
    if key_for_used:
        county_sheet_df = county_sheet_df.loc[county_sheet_df[key_for_used].notnull()]

    # now filter to take only rows where our url field is not blank
    county_sheet_df = county_sheet_df.loc[county_sheet_df[key_containing_url].notnull()]

    # get df of rows with unique county and url (take only the first url/county combination if multiple. Should be ok
    # cause multiple urls should be the generic)
    county_sheet_df_unique_url = county_sheet_df.drop_duplicates(
        subset=key_containing_url, keep='first')

    for _, row in county_sheet_df_unique_url.iterrows():
        url = row[key_containing_url]
        if 'http' not in url:
            url = f"https://{url}"

        if url_error(url) is None:
            logger.info(f"OK: '{url}'")
        else:
            logger.warning(f"Error with:'{url}', '{url_error(url)}'")

        if open_browser == 'all':
            webbrowser.open(url, new=2)
            if url_error(url) is not None:
                webbrowser.open(f"https://www.google.com/search?q={row['{state}']} {row['{cntytoprint}']} "
                            f"board of elections", new=2)
        elif open_browser == 'error' and url_error(url) is not None:
            webbrowser.open(url, new=2)
            webbrowser.open(f"https://www.google.com/search?q={row['{state}']} {row['{cntytoprint}']} "
                            f"board of elections", new=2)


def check_short_boe_urls(key_containing_short_url: str = None, key_containing_url: str = None,
                         key_for_used: str = None, open_browser=None, boe_xls=None) -> None:
    """Verifies that short URLs in a BOE xlsx redirect to their expected final URLs.

    Compares the resolved destination of each short URL against the corresponding
    full URL column and logs mismatches as errors.

    Args:
        key_containing_short_url (str): Column key holding the short URL
            (e.g. '{shorturl}'). Prompts if None.
        key_containing_url (str): Column key holding the expected final URL
            (e.g. '{url}'). Prompts if None.
        key_for_used (str): Column key identifying active county rows
            (e.g. '{priority}'). Prompts if None.
        open_browser: Browser opening mode passed to get_browser_option. Prompts if None.
        boe_xls: Path to the BOE xlsx file. Prompts via file picker if None.
    """

    from .read_boe_xls import read_boe_xls

    if boe_xls is None:
        # boe_xls = get_file_name("Pick BOE file (xlsx) with urls to check.",
        #                          initial_dir=ATTACHMENTS_INITIAL_DIR,
        #                          title2="Pick BOE file (xlsx)")

        if True:  # not hardcoded
            boe_xls = select_file(
                title="Pick BOE file (xlsx) with urls to check.",
                start_dir=ATTACHMENTS_INITIAL_DIR,
                files_like="*.xlsx",
                choices=["Select", "Cancel"],
                mode="file",  # file, dir or both
                title2="Will be in the INPUT subdirectory"
            )
        else:
            boe_xls = '/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/Attachments/Campaigns/Test box and image/Input/BOE Info FL.xlsx'

    logger.info(f"BOE xls being checked is: '{boe_xls}")

    county_sheet_keys, county_sheet_df = read_boe_xls(boe_xls)

    # for comparison url get key like '{url}' if not passed
    if key_containing_url is None:
        # key_containing_url = select_from_list(county_sheet_keys, 'Select tag for url field to match',
        #                                       select_type='radio')
        key_containing_url = list_pick(county_sheet_keys,
                                     'Select tag for FINAL URL',
                                     "Select tag for url field to match, often '{url}'",
                                     select_mode='single',
                                     pre_select=False,
                                        allow_none=False,
                                       )[0]
    # make sure it's on file
    key_containing_url = key_containing_url.lower()
    logger.info(f"Key field being checked is '{key_containing_url}'")
    # make sure key is on file
    if key_containing_url not in county_sheet_keys:
        exit_yes(f"Url key '{key_containing_url}' not in county keys: \n\n{', '.join(county_sheet_keys)}")

    # for short url get key like '{bitlyurl}' if not passed
    if key_containing_short_url is None:
        # key_containing_short_url = select_from_list(county_sheet_keys, 'Select tag for short url field to be checked',
        #                                             select_type='radio')
        key_containing_short_url = list_pick(county_sheet_keys,
                                             'Select tag for SHORT URL',
                                             "Select tag for short url field to be checked (usually '{shorturl}' or "
                                             "'{bitlyurl}')",
                                             select_mode='single',
                                             pre_select=False,
                                             allow_none=False,
                                             )[0]
    # make sure it's on file
    key_containing_short_url = key_containing_short_url.lower()
    logger.info(f"Key field being checked is '{key_containing_short_url}'")
    # make sure key is on file
    if key_containing_short_url not in county_sheet_keys:
        exit_yes(f"Url key '{key_containing_short_url}' not in county keys: \n\n{', '.join(county_sheet_keys)}")

    # get key like '{priority}' if not passed
    if key_for_used is None:
        # key_for_used = select_from_list(county_sheet_keys,
        #                                 'Select tag to id counties being used (None for all)',
        #                                 select_type='radio')
        key_for_used = list_pick(county_sheet_keys,
                                             'Select tag',
                                             "Select tag to id counties being used (often '{priority},'None for all "
                                             "counties)",
                                             select_mode='single',
                                             pre_select=False,
                                            allow_none=True,
                                             )[0]
    # make sure key is on file
    if key_for_used is not None and key_for_used is not '':  # None ok from select = all so have to check
        key_for_used = key_for_used.lower()
        if key_for_used not in county_sheet_keys:
            exit_yes(f"Key 'for used' '{key_for_used}' not in county keys: \n\n{', '.join(county_sheet_keys)}")
    logger.info(f"Key field being used to limit counties is '{key_for_used}'")

    # if a field is specified to identify used rows (those we are interested in) then filter the df
    if key_for_used is not None and key_for_used is not '':
        county_sheet_df = county_sheet_df.loc[county_sheet_df[key_for_used].notnull()]

    # only rows where our url field is not blank
    county_sheet_df = county_sheet_df.loc[county_sheet_df[key_containing_short_url].notnull()]

    # get df of rows with unique county and url (take only the first url/county combination if multiple. Should be ok
    # cause multiple urls should be the generic)
    county_sheet_df_unique_url = county_sheet_df.drop_duplicates(
        subset=key_containing_short_url, keep='first')

    # compare final redirected destination of short url and ur and list if different
    for _, row in county_sheet_df_unique_url.iterrows():
        short_url = row[key_containing_short_url]
        url = row[key_containing_url]
        if 'http' not in short_url:
            short_url = f"https://{short_url}"
        if 'http' not in url:
            url = f"https://{url}"
        if final_url(short_url) == final_url(url):
            logger.info(f"'{short_url}' Ok. It goes to: '{final_url(short_url)}'")
        else:
            logger.error("")
            logger.error(f"Mismatch: priority '{row['{priority}']}', '{short_url}' goes to: "
                         f"{'ERROR' if url_error(short_url) is not None else ''} '{final_url(short_url)}'")
            # logger.error(f"Mismatch: priority '{row['{priority}']}', '{short_url}' comparison is:'{final_url(url)}'")
            logger.error(f"Comparison is:'{final_url(url)}'")
            logger.error("")
            a = 1
    a = 1


if __name__ == '__main__':
    logger = setup_loguru("DEBUG", "DEBUG")

    # url_error("https://www.fcva.us/departments/office-of-elections-voter-registration")
    # check_boe_urls(key_containing_url='{xxx}', key_for_used='{priority}')
    # check_boe_urls(key_containing_url='{evurl}', key_for_used='{xxx}')
    # check_boe_urls(key_containing_url='{evurl}', key_for_used='{priority}')
    # check_boe_urls(key_containing_url='{evurl}', key_for_used='{priority}', open_browser=False,)
    # check_boe_urls(key_containing_url='{evurl}', key_for_used='{priority}', open_browser=False,
    #                boe_xls=Path("~/Dropbox/Postcard Files/PythonPrograms/Development/LeviTool/Work/TEST BOE Info VA.xlsx").expanduser(),)
    # check_boe_urls(key_containing_url='{url}', key_for_used='{priority}', open_browser=True,
    #                boe_xls=Path("~/Dropbox/Postcard Files/Attachments/Campaigns/VA Primary 1-2024/Input/BOE Info VA.xlsx").expanduser(),)
    # check_boe_urls(key_containing_url=None, key_for_used='all', open_browser=False,
    #                boe_xls=Path("~/Dropbox/Postcard Files/Attachments/Campaigns/VA Primary 1-2024/Input/BOE Info VA.xlsx").expanduser(),)
    # check_short_boe_urls(key_containing_short_url=None, key_containing_url=None, key_for_used='all', open_browser=False,
    #                boe_xls=Path("~/Dropbox/Postcard Files/PythonPrograms/Development/LeviTool/Work/TEST BOE Info VA.xlsx").expanduser(),)
    check_boe_urls(key_containing_url='{specificurl}', key_for_used='{use}', open_browser='All',
                   boe_xls=Path("~/Dropbox/Postcard Files/PythonPrograms/Development/LeviTool/Work/TEST BOE Info VA.xlsx").expanduser())
    # check_short_boe_urls(open_browser=True)
