"""this file the translated data from an excel file and moves it
to it's respective place in a .ts file
"""

import getopt
import re
import sys
import xml.etree.ElementTree as xml

import openpyxl
from googletrans import Translator
from openpyxl.workbook.workbook import Workbook


class Translation:
    """This holds the information gathered from the Excel file"""

    def __init__(self, source: str, translation: str, comments: str, identifier: int):
        self.source = source
        self.translation = translation
        self.comments = comments
        self.id = identifier
        self.verification_text = ""


def get_args() -> list:
    """This grabs the command line arguments parsed in alongside the file
    TODO: make it work without command line args

    Returns:
        str, bool:
        excel_file is the filepath to the excel file,
        verification_mode is a boolean that tells weather or not
        we want to output a verification file
    """

    # see if we gave any command line arguments
    cmd_args = None
    try:
        cmd_args = sys.argv[1:]

    # if there's an unexpected error, just print it
    except Exception as error:
        print(error)

    # if there are command line arguments,
    verification_mode = False
    excel_file = None
    ts_file = None
    if cmd_args:

        # try to get them
        try:
            opts, _ = getopt.getopt(cmd_args, "hvi:o:")

        # if we can't, show the usage and close
        except getopt.GetoptError:
            print("<filename>.py -i <input_file_path> -o <output_file_path>")
            print('optionally you can use "-v" to enable verification output')
            sys.exit(2)

        # after getting the arguments, loop through them all
        for opt, arg in opts:

            # if we have the "help" argument
            #   explain the usage of the command line arguments
            if opt == "-h":
                print("usage: <filename>.py -i <input_file_path> -o <output_file_name>")
                print('optionally you can use "-v" to enable verification output')
                sys.exit()

            # if we have the "input_file" argument,
            #   set the correct filepath variable
            elif opt == "-i":
                excel_file = arg

            # if we have the "output_file" argument,
            #   set the correct filepath variable
            elif opt == "-o":
                ts_file = arg

            # if we have the "verification" argument,
            #   set the correct verification variable
            elif opt == "-v":
                verification_mode = True

    # if there wasn't any command line arguments,
    else:
        print("<filename>.py -i <input_file_path> -o <output_file_path>")
        print('optionally you can use "-v" to enable verification output')

    return [excel_file, ts_file, verification_mode]


def get_data(filepath: str) -> list:
    """this will take in a filepath to an Excel file and grab the data from it
    and put it into a Translation class list

    Args:
        filepath (str): This should be a filepath to an Excel file

    Returns:
        list: This will be the list of all the Translation classes that were
        created, sorted alphabetically by their source text
    """

    # try to load the Excel file located at the filepath provided
    #   in read only mode
    try:
        input_data = openpyxl.load_workbook(filename=filepath, read_only=True)

    # if that fails, print the error
    except Exception as error:
        print(error)
        return

    # get the current active Sheet in the Excel file
    active_sheet = input_data.active

    # build the final datalist out of Translation classes in the format:
    # Translation.source :=
    #   the_currently_active_sheet[Column: A, Row: loop_index]

    # Translation.translation :=
    #   the_currently_active_sheet[Column: B, Row: loop_index]

    # Translation.comments :=
    #   the_currently_active_sheet[Column: C, Row: loop_index]

    # Translation.id :=
    #   loop_index
    # ! This never ends up being used but is stored for possible future use
    data = [
        Translation(
            active_sheet["A{}".format(i)].value,
            active_sheet["B{}".format(i)].value,
            active_sheet["D{}".format(i)].value,
            i,
        )
        for i in range(2, active_sheet.max_row + 1)
    ]

    # close the Excel file as loading it in read only mode requires closing
    input_data.close()

    # sort the Translation class list alphabetically by source text
    data.sort(key=lambda x: (x.source or "").lower())

    return data


def get_translations(translation_class_list: list) -> list:
    """this takes in a list of Translation classes and returns a list
    containing just the "translation" variable from them

    Args:
        translation_class_list (list): A list of Translation classes

    Returns:
        list: A list containing just the "translation" variables
        from the original list
    """
    return [i.translation for i in translation_class_list if i.source is not None]


def html_prep(input_string: str) -> str:
    """The Excel file uses characters that aren't allowed in the .ts file.
    This will take in an input string and replace
    all the nonvalid characters with valid ones

    Args:
        input_string (str): The string to be cleansed of it's sins

    Returns:
        string: the cleaned string
    """
    return (
        input_string.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
        .replace("\xa0", "&#160;")
    )


def cleanup_translation(input_string: str) -> str:
    """This takes in a string and cleans out all the escape characters
    for use in more accurate translations

    Args:
        input_string (str): the string to be cleaned

    Returns:
        string: the cleaned string
    """
    # search for any html tags
    html_tag = re.findall("<[^>]*>", input_string)

    # if there's any, loop through them and convert them into a space
    if html_tag:
        for sequence in html_tag:
            input_string = input_string.replace(sequence, " ")

    # fix the common escape sequences
    input_string = input_string.replace("\n", " ").replace("\xa0", " ")

    # get rid of the "%" references
    for i in range(1, 11):
        input_string = input_string.replace("%{}".format(i), " ")

    return input_string


def translate_translations(translations_list: list) -> list:
    """this takes in a list of strings to be translated and translates them

    Args:
        translations_list (list): a list of strings to be translated

    Returns:
        list: a list of the translated strings
    """
    translations = Translator().translate(translations_list, dest="english")
    return [translation.text for translation in translations]


def verification_file(translation_class_list: list, language: str):
    """this creates an Excel file for organization of the verification data

    Args:
        translation_class_list (list): the list of the Translation classes
        language (str): the language that we translated from
    """

    # create an empty Excel file, name the sheet and select the current sheet
    output_workbook = Workbook()
    output_workbook.title = "Verification Data"
    output_worksheet = output_workbook.active

    # load the correct amount of rows and columns
    for x in range(1, 3):
        for y in range(1, len(translation_class_list) + 1):
            output_worksheet.cell(row=x, column=y)

    # label the columns
    output_worksheet["A1"] = "Source"
    output_worksheet["B1"] = "Translated Translation"

    # put the data in the Excel file with the original source text in the
    #   A column and the translated translation in the B column
    index = 2
    for translation in translation_class_list:
        if translation.source is not None:
            output_worksheet["A{}".format(index)] = translation.source
            output_worksheet["B{}".format(index)] = translation.verification_text
            index += 1

    # save the file
    output_workbook.save("Verification_Data_{}.xlsx".format(language))


def write_to_ts_file(filepath: str, translation_class_list: list):
    """this takes the input .ts filepath and puts the Excell data
    into its coresponding location

    Args:
        filepath (str): the filepath to the .ts file
        translation_class_list (list): thi list of Translation classes
    """

    # Load the .ts file and make a variable for it's root
    input_file = xml.ElementTree(file=filepath)
    input_file_root = input_file.getroot()

    # loop through all the <message> tags and find the Translation class with
    # the coresponding source text.
    # Then put the translation into the <translation> tag
    for message in input_file_root.iter("message"):
        source = message.find("source")
        for translation in translation_class_list:
            if translation.source == source.text:
                message.find("translation").text = translation.translation
                break

    # save the file
    input_file.write(filepath)


def main():
    """this is the main function that handles
    the ordering of the other functions"""

    # we first get the file path of the Excel file, .ts file, and
    #   wether or not we want a verification file
    excel_file, ts_file, verification_mode = get_args()

    # next, we get a list of Translation classes containing
    #   the translations in the excel file.
    # we also get the current language from the end of the file
    #   for use in locating the coresponding .ts file
    translation_data = get_data(excel_file) if excel_file else None
    lang = excel_file[-7:-5]

    # next we see if that list even exists
    if translation_data:

        # if it does, first we see if the user wants a verification file
        if verification_mode:

            # if they do, First we get all the translations from
            #   the list of Translation classes
            foreign_words = get_translations(translation_data)

            # next we strip out all the escape characters to make the process
            #   of translating more accurate
            foreign_words = [
                cleanup_translation(translation) for translation in foreign_words
            ]

            # now we translate the translations back to english
            foreign_words = translate_translations(foreign_words)

            # next we loop through the Translation class list and
            #   add the translated translations into the class's
            #   verification_text variable
            index = 0
            for translation in translation_data:
                if translation.source is not None:
                    translation.verification_text = foreign_words[index]
                    index += 1

            # Finally we write the data to a verification Excel file
            verification_file(translation_data, lang)

        # next we need to make sure that the text fits within the constraints
        #   of the .ts file
        # we do this by replacing the symbols in the text to
        #   HTML compliant ones
        for translation in translation_data:
            if translation.translation:
                translation.translation = html_prep(translation.translation)

        # finally we write the data from the excel file to the .ts file
        write_to_ts_file(ts_file, translation_data)


if __name__ == "__main__":
    main()
