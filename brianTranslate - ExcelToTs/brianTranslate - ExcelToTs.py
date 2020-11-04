"""this file the translated data from an excel file and moves it to it's respective place in a .ts file
"""

import getopt
import re
import sys

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


class tag:
    """This holds information gathered from the .ts file"""

    def __init__(
        self, typ: str, content_start: int, content_end: int, content: str, index: int
    ):
        self.typ = typ
        self.content_start = content_start
        self.content_end = content_end
        self.content = content
        self.index = index

    def __str__(self):
        return self.content


def get_args():
    """This grabs the command line arguments parsed in alongside the file
    TODO: make it work without command line args

    Returns:
        str, bool:
        excel_file is the filepath to the excel file,
        verification_mode is a boolean that tells weather or not we want to output a verification file
    """

    # see if we gave any command line arguments
    try:
        cmd_args = sys.argv[1:]

    # if we didn't make the variable equal to None
    except IndexError:
        cmd_args = None

    # if there's some other error, just print it
    except Exception as error:
        print(error)

    # if there are command line arguments,
    verification_mode = False
    excel_file = None
    if cmd_args:

        # try to get them
        try:
            opts, _ = getopt.getopt(cmd_args, "hi:v")

        # if we can't, show the usage and close
        except getopt.GetoptError:
            print("brianTranslate - ExcelToTs.py -i <input_file_path>")
            print(
                'optionally you can use "-v" after the input file to enable verification output'
            )
            sys.exit(2)

        # after getting the arguments, loop through them all
        for opt, arg in opts:

            # if we have the "help" argument, explain the usage of the command line arguments
            if opt == "-h":
                print("usage: brianTranslate - ExcelToTs.py -i <input_file_path>")
                print(
                    'optionally you can use "-v" after the input file to enable verification output'
                )
                sys.exit()

            # if we have the "input_file" argument, set the correct filepath variable
            elif opt == "-i":
                excel_file = arg

            # if we have the "verification" argument, set the correct verification variable
            elif opt == "-v":
                verification_mode = True

    # if there wasn't any command line arguments,
    else:
        print("brianTranslate - ExcelToTs.py -i <input_file>")
        print(
            'optionally you can use "-v" after the input file to enable verification output'
        )

    return excel_file, verification_mode


def get_data(filepath: str):
    """this will take in a filepath to an Excel file and grab the data from it and put it into a Translation class list

    Args:
        filepath (str): This should be a filepath to an Excel file

    Returns:
        list: This will be the list of all the Translation classes that were created, sorted alphabetically by their source text
    """

    # try to load the Excel file located at the filepath provided in read only mode
    try:
        input_data = openpyxl.load_workbook(filename=filepath, read_only=True)

    # if that fails, print the error
    except Exception as error:
        print(error)
        return

    # get the current active Sheet in the Excel file
    active_sheet = input_data.active

    # build the final datalist out of Translation classes in the format:
    # Translation.source      := the_currently_active_sheet[Column: A, Row: loop_index]
    # Translation.translation := the_currently_active_sheet[Column: B, Row: loop_index]
    # Translation.comments    := the_currently_active_sheet[Column: C, Row: loop_index]
    # Translation.id          := loop_index  <- #! This never ends up being used but is stored for possible future use
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


def get_translations(translation_class_list: list):
    """this takes in a list of Translation classes and returns a list containing just the "translation" variable from them

    Args:
        translation_class_list (list): A list of Translation classes

    Returns:
        list: A list containing just the "translation" variables from the original list
    """
    return [i.translation for i in translation_class_list if i.source != None]


def html_prep(input_string: str):
    """The Excel file uses characters that aren't allowed in the .ts file.
    This will take in an input string and replace all the nonvalid characters with valid ones

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


def cleanup_translation(input_string: str):
    """This takes in a string and cleans out all the escape characters for use in more accurate translations

    Args:
        input_string (str): the string to be cleaned

    Returns:
        string: the cleaned string
    """
    # search for any html tags
    html_tag = re.findall("<[^>]*>", input_string)

    # if there's any,
    if html_tag:

        # loop through them and convert them into a space
        for sequence in html_tag:
            input_string = input_string.replace(sequence, " ")

    # fix the common escape sequences
    input_string = input_string.replace("\n", " ").replace("\xa0", " ")

    # get rid of the "%" references
    for i in range(1, 11):
        input_string = input_string.replace("%{}".format(i), " ")

    return input_string


def translate_translations(translations_list: list):
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

    # put the data in the Excel file with the original source text in the A column and the translated translation in the B column
    index = 2
    for translation in translation_class_list:
        if translation.source != None:
            output_worksheet["A{}".format(index)] = translation.source
            output_worksheet["B{}".format(index)] = translation.verification_text
            index += 1

    # save the file
    output_workbook.save("Verification_Data_{}.xlsx".format(language))


def get_tag_locations(language: str):
    """This grabs the location and the contents of each source and translation tag in the .ts file

    Args:
        language (str): the language of the input Excel file, used to find the coresponding .ts file

    Returns:
        list: the list of Tag classes sorted alphabetically by their source text
    """

    # setup the empty final list
    extracted_tags = []

    # try to open the coresponding .ts file
    try:
        with open(
            "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
            "r",
            encoding="utf-8",
        ) as input_file:

            # get a list of the raw data of the file, seperated by lines
            raw_data = open(
                "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
                "r",
                encoding="utf-8",
            ).readlines()

            # iterate through the lines of the file
            for line_index, line in enumerate(input_file):

                # find the specific tags
                if "<source" in line or "<translation" in line:

                    # set these up outside class definition for use in class definition
                    content_start = line.find(">") + 1
                    content_end = line.find("<", content_start)

                    # check to see if we even need to do lookahead
                    if "</source" not in line and "</translation" not in line:

                        # if we do, setup a few lines of lookahead
                        #! this looks 9 lines forward in the data array for the closing tag,
                        #! if the data you're translating is longer than that, change this variable
                        lines_of_lookahead = 9

                        lookahead = raw_data[
                            line_index : line_index + lines_of_lookahead + 1
                        ]

                        # loop through the lookahead
                        for line_ahead_number, line_ahead in enumerate(lookahead):

                            # find the end tag
                            if (
                                "</source" in line_ahead
                                or "</translation" in line_ahead
                            ):

                                # merge the range of lines and cut out the content
                                content_range = "".join(
                                    raw_data[
                                        line_index : line_ahead_number + line_index + 1
                                    ]
                                )
                                content_start = content_range.find(">") + 1
                                content_end = content_range.find("<", content_start)
                                content_range = content_range[content_start:content_end]
                                break
                    # if we don't need to do lookahead, just get the content range
                    else:
                        content_range = line[content_start:content_end]

                    # add a Tag class to the list
                    extracted_tags.append(
                        tag(
                            line[line.find("<") + 1 : line.find(">")],
                            content_start,
                            content_end,
                            content_range,
                            line_index,
                        )
                    )

    # if we can't find the file, print out what the program was expecting
    except FileNotFoundError:
        print(
            'file was not found, Expected file in this directory to be named "exspiron2Xi_{}_{}"'.format(
                language.lower(), language.upper()
            )
        )

    # reformat the tags list for easier use later
    # tags list is reformatted as such:
    # tags_list := [[source_tag, it's_translation], etc...]
    extracted_tags = [
        [tag, extracted_tags[index + 1]]
        for index, tag in enumerate(extracted_tags)
        if tag.typ == "source"
    ]

    # sort the new tags list alphabetically by the source text
    extracted_tags.sort(key=lambda x: str(x[0]).lower())

    return extracted_tags


def put_data_into_file(translation_class_list, locations, language: str):
    """this gets the data gathered from the Excel file and gets it formatted to be ready to write to it's coresponding .ts file

    Args:
        translation_class_list (list): the list of all the translation classes
        locations ([type]): the list of locations in the .ts file where the translations should go
        language (str): the language we are dealing in

    Returns:
        list: a list of lines ready to be written to the .ts file
    """

    # get a list of the raw data of the file, seperated by lines
    raw_data = open(
        "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
        "r",
        encoding="utf-8",
    ).readlines()

    # loop through all the Excel file's translations and put them into their coresponding translation in the .ts file data list
    # once we do that we then get rid of the compared source text in the locations list
    for location in locations:
        for translation in translation_class_list:
            if translation.source == location[0].content:
                location[1].content = translation.translation
                del location[0]
                break

    # turn the list of lists of single items into just one list
    holder = []
    for location in locations:
        for tag in location:
            holder.append(tag)
    locations = holder

    # add the translation to the raw data list and mark the translation as "complete"
    # marking it as "complete" means changeing "<translation type="unfinished">" to "<translation>"
    # if there was no translation, we don't do this
    for location in locations:
        raw_data[location.index] = (
            raw_data[location.index][: location.content_start]
            + location.content
            + raw_data[location.index][location.content_end :]
        )
        content_start = raw_data[location.index].find(">") + 1
        content_end = raw_data[location.index].find("<", content_start)
        if content_start != content_end:
            raw_data[location.index] = raw_data[location.index].replace(
                'translation type="unfinished"', "translation"
            )

    return raw_data


def write_to_ts_file(input_list: list, language: str):
    """this writes a raw data list to a .ts file

    Args:
        input_list (list): the raw data list
        language (str): the language, used for finding the coresponding .ts file
    """
    print("Transferring to the TS file...")
    with open(
        "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
        "w",
        encoding="utf-8",
    ) as input_file:
        input_file.writelines(input_list)
    print("Excel transfer Complete!")


def main():
    """this is the main function that handles the ordering of the other functions"""

    # we first get the file path of the Excel file and wether or not we want a verification file
    excel_file, verification_mode = get_args()

    # next, we get a list of Translation classes containing the translations in the excel file.
    # we also get the current language from the end of the file for use in locating the coresponding .ts file
    translation_data = get_data(excel_file) if excel_file else None
    lang = excel_file[-7:-5]

    # next we see if that list even exists
    if translation_data:

        # if it does, first we see if the user wants a verification file
        if verification_mode:

            # if they do, First we get all the translations from the list of Translation classes
            foreign_words = get_translations(translation_data)

            # next we strip out all the escape characters to make the process of translating more accurate
            foreign_words = [
                cleanup_translation(translation) for translation in foreign_words
            ]

            # now we translate the translations back to english
            foreign_words = translate_translations(foreign_words)

            # next we loop through the Translation class list and add the translated translations into the class's verification_text variable
            index = 0
            for translation in translation_data:
                if translation.source != None:
                    translation.verification_text = foreign_words[index]
                    index += 1

            # Finally we write the data to a verification Excel file
            verification_file(translation_data, lang)

        # next we get the locations of the translations in the .ts file
        tag_locations = get_tag_locations(lang)

        # next we need to make sure that the text fits within the constraints of the .ts file
        # we do this by replacing the symbols in the text to HTML compliant ones
        for translation in translation_data:
            if translation.translation:
                translation.translation = html_prep(translation.translation)

        # now we get the raw data needed for writing to the .ts file
        final_output = put_data_into_file(translation_data, tag_locations, lang)

        # finally we write the raw data to the .ts file
        write_to_ts_file(final_output, lang)


if __name__ == "__main__":
    main()
