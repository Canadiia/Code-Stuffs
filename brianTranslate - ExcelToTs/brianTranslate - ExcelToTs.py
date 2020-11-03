    """this file the translated data from an excel file and moves it to it's respective place in a .ts file
    TODO: DOCUMENTATION
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
        str, bool: excel_file is the filepath to the excel file,
        verification_mode is a boolean that tells weather or not we want to output a verification file
    """

    # see if we gave any command line arguments
    try:
        cmd_args = sys.argv[1:]

    # if we didn't make the variable None
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
            print("brianTranslate - ExcelToTs.py -i <input_file>")
            print(
                'optionally you can use "-v" after the input file to enable verification output'
            )
            sys.exit(2)

        # loop through all the arguments
        for opt, arg in opts:

            # if we have the "help" argument, explain the usage of the command line arguments
            if opt == "-h":
                print("usage: brianTranslate - ExcelToTs.py -i <input_file>")
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
    try:
        input_data = openpyxl.load_workbook(filename=filepath, read_only=True)
    except Exception as error:
        print(error)
        return
    length = input_data.active
    data = [
        Translation(
            length["A{}".format(i)].value,
            length["B{}".format(i)].value,
            length["D{}".format(i)].value,
            i,
        )
        for i in range(2, length.max_row + 1)
    ]
    input_data.close()
    data.sort(key=lambda x: (x.source or "").lower())
    return data


def get_translations(translation_class_list: list):
    return [i.translation for i in translation_class_list if i.source != None]


def html_prep(input_string: str):
    return (
        input_string.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
        .replace("\xa0", "&nbsp;")
    )


def cleanup_translations(translation_class_list: list):
    holder = []
    for translation in translation_class_list:
        final_string = translation

        # search for any html tags
        html_tag = re.findall("<[^>]*>", final_string)

        # if there's any,
        if html_tag:

            # loop through them
            for sequence in html_tag:

                # convert them into a space
                final_string = final_string.replace(sequence, " ")

        # fix the common escape sequences
        final_string = final_string.replace("\n", " ").replace("\xa0", " ")

        for i in range(1, 11):
            final_string = final_string.replace("%{}".format(i), " ")

        holder.append(final_string)

    return holder


def translate_translations(translations_list: list):
    translations = Translator().translate(translations_list, dest="english")
    return [translation.text for translation in translations]


def verification_file(translation_class_list: list, language: str):
    output_workbook = Workbook()
    output_workbook.title = "Verification Data"
    output_worksheet = output_workbook.active
    for x in range(1, 3):
        for y in range(1, len(translation_class_list) + 1):
            output_worksheet.cell(row=x, column=y)

    output_worksheet["A1"] = "Source"
    output_worksheet["B1"] = "Translated Translation"

    index = 2
    for translation in translation_class_list:
        if translation.source != None:
            output_worksheet["A{}".format(index)] = translation.source
            output_worksheet["B{}".format(index)] = translation.verification_text
            index += 1

    output_workbook.save("Verification_Data_{}.xlsx".format(language))


def get_tag_locations(language: str):
    extracted_tags = []
    try:
        # open the file
        with open(
            "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
            "r",
            encoding="utf-8",
        ) as input_file:

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
                        #! this looks 9 lines forward in the data array for the closeing tag,
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
                    else:
                        content_range = line[content_start:content_end]

                    extracted_tags.append(
                        tag(
                            line[line.find("<") + 1 : line.find(">")],
                            content_start,
                            content_end,
                            content_range,
                            line_index,
                        )
                    )

    except FileNotFoundError:
        print(
            'file was not found, Expected file in this directory to be named "exspiron2Xi_{}_{}"'.format(
                language.lower(), language.upper()
            )
        )
    extracted_tags = [
        [tag, extracted_tags[index + 1]]
        for index, tag in enumerate(extracted_tags)
        if tag.typ == "source"
    ]

    extracted_tags.sort(key=lambda x: str(x[0]).lower())
    return extracted_tags


def put_data_into_file(translation_class_list, locations, language: str):
    raw_data = open(
        "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
        "r",
        encoding="utf-8",
    ).readlines()

    for location in locations:
        for translation in translation_class_list:
            if translation.source == location[0].content:
                location[1].content = translation.translation
                del location[0]
                break

    holder = []
    for location in locations:
        for tag in location:
            holder.append(tag)
    locations = holder

    for location in locations:
        raw_data[location.index] = (
            raw_data[location.index][: location.content_start]
            + location.content
            + raw_data[location.index][location.content_end :]
        ).replace('translation type="unfinished"', "translation")

    return raw_data


def write_to_ts_file(input_list: list, language: str):
    print("Transferring to the TS file...")
    with open(
        "exspiron2Xi_{}_{}.ts".format(language.lower(), language.upper()),
        "w",
        encoding="utf-8",
    ) as input_file:
        input_file.writelines(input_list)
    print("Excel transfer Complete!")


def main():

    # explication
    excel_file, verification_mode = get_args()

    # explication
    translation_data = get_data(excel_file) if excel_file else None
    lang = excel_file[-7:-5]

    # explication
    if translation_data:

        # explication
        if verification_mode:

            # explication
            foreign_words = get_translations(translation_data)

            # explication
            foreign_words = cleanup_translations(foreign_words)

            # explication
            foreign_words = translate_translations(foreign_words)

            # explication
            index = 0
            for translation in translation_data:

                # explication
                if translation.source != None:
                    translation.verification_text = foreign_words[index]
                    index += 1

            # explication
            verification_file(translation_data, lang)

        # explication
        tag_locations = get_tag_locations(lang)

        # explication
        for index, translation in enumerate(translation_data):
            if translation.translation:
                translation.translation = html_prep(translation.translation)

        # explanation
        final_output = put_data_into_file(translation_data, tag_locations, lang)

        write_to_ts_file(final_output, lang)

        input()


if __name__ == "__main__":
    main()
