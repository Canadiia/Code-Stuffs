#!/usr/bin/python3

'''This script take translations from an excel file and put them into a Qt .ts file.
See the `print_usage_message()` function for how to invoke this script.

For the format that the Excel file needs to be in, see the `read_translation_records`
function.
'''

import sys
from typing import List, NamedTuple, Optional
import openpyxl
from lxml import etree


# == Data Containers ==

class Arguments(NamedTuple):
    '''Command line arguments'''
    excel_filepath: str
    ts_filepath: str

class TranslationRecord(NamedTuple):
    '''This holds the information gathered from the Excel file'''
    english_source: str
    translation: str


# == Functions ==

def print_usage_message() -> None:
    '''Print a message on how to use this program'''
    print('Usage:')
    print('  xlsx_2_ts.py <translation_source>.xlsx <destination>.ts')


def validate_arguments(raw_args: List[str]) -> Optional[Arguments]:
    '''Validates if the arguments provided at the command line are good.  If so
    returns Those arguments.  If not, prints the usage message and returns None'''

    # First check count
    argc = len(raw_args)
    if argc < 3:
        print('Not enough arguments.')
        print_usage_message()
        return None

    # Next, check extensions
    args = Arguments(
        excel_filepath=raw_args[1],
        ts_filepath=raw_args[2],
    )

    # Validate
    if not args.excel_filepath.lower().endswith('.xlsx'):
        print('First argument must be to an Excel file.')
        print_usage_message()
        return None

    if not args.ts_filepath.lower().endswith('.ts'):
        print('Second argument must be to a Qt .ts XML file.')
        print_usage_message()
        return None

    # Note: A more proper validating would require reading the files in, (maybe just their headers)
    #       and then verifying they actually are excel files (and a .ts XML).  But here we're relying
    #       on correct extensions

    # Seems all good
    return args


def read_translation_records(excel_filepath: str) -> List[TranslationRecord]:
    '''This will read a provided Excel file and then parse the rows/cells into
    a list of translation records.

    The Excel files follow the format where the first column (A) is the original
    source English.  The second column (B) is the translation into the target
    language.

    The third column (C) is a possible abbreviation for the translation, where as
    the fourth (D) are notes that we provided to the translators.  These are not
    used by us, so we ignore them for now.

    Translation records start with the third row'''

    # There should be only one sheet; it's active and where the data is
    excel_file = openpyxl.load_workbook(excel_filepath, read_only=True)
    sheet = excel_file.active

    # Note: would prefer to use `with` syntax for file opening/closing, but it
    #       doesn't seem to be available with openpyxl

    records = []

    # Reaching in a batch of rows (and columns) is much faster
    upper_limit = sheet.max_row + 1
    rows = sheet['A3:B%i' % upper_limit]

    # Go throuh each row (a two element tuple in our case)
    for (english_cell, translation_cell) in rows:
        # Pull their values
        ev = english_cell.value
        tv = translation_cell.value

        # Store? Skip? Error?
        if (not ev) and (not tv):
            # If both cells were empty, skip
            continue
        elif ev and tv:
            # If both cells have entries, store it
            records.append(TranslationRecord(
                english_source=ev,
                translation=tv
            ))
        else:
            # This is bad, one of them is None and the other has a value, ERROR
            raise Exception((
                'Mismatching English and Translation; one cel is empty',
                ev,
                tv
            ))

    # Cleanup and return Translations
    excel_file.close()
    return records


def main() -> None:
    '''Main routine of the program'''

    # First validate the arguments
    args = validate_arguments(sys.argv)
    if not args:
        print('Error; arguments incorrect. Exiting.', file=sys.stderr)
        sys.exit(1)

    # Read in the Excel file
    recs = read_translation_records(args.excel_filepath)

    # Read in the .ts XML
    ts_xml = etree.parse(args.ts_filepath)
    all_souce_tags = ts_xml.findall('.//source')        # find all of the <source> tags using a recursive search

    # Go through each translation entry
    for rec in recs:
        # Find the <source> tags that have the matching english
        matching_tags = filter(lambda x: (x.text == rec.english_source), all_souce_tags)

        # For each of the matching ones, edit the <translation> tag to use the translation
        for tag in matching_tags:
            message_tag = tag.getparent()
            message_tag.find('translation').text = rec.translation

    # All done, now save it
    ts_xml.write(args.ts_filepath, pretty_print=True, xml_declaration=True, encoding='utf-8')


if __name__ == '__main__':
    main()
