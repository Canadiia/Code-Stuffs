import getopt
import re
import sys

import openpyxl
from googletrans import Translator
from openpyxl.workbook.workbook import Workbook


class Translation:
    def __init__(self, source, translation, comments, identifier):
        self.source = source
        self.translation = translation
        self.comments = comments
        self.id = identifier
        self.verification_text = ""


class tag:
    def __init__(
        self,
        tagStart: int,
        typ: str,
        index: int,
        contentStart: int,
        contentEnd: int,
        content: str,
    ):
        self.tagStart = tagStart
        self.typ = typ
        self.index = (index, tagStart)
        self.contentStart = contentStart
        self.contentEnd = contentEnd
        self.content = content


def get_args():
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
            opts, args = getopt.getopt(cmd_args, "hi:v")

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


def get_data(input_file):
    data = []
    try:
        input_data = openpyxl.load_workbook(filename=input_file, read_only=True)
    except Exception as error:
        print(error)
        return
    length = input_data.active
    for i in range(2, length.max_row + 1):
        data.append(
            Translation(
                length["A{}".format(i)].value,
                length["B{}".format(i)].value,
                length["D{}".format(i)].value,
                i,
            )
        )
    input_data.close()
    return data


def get_translations(input_data):
    return [i.translation for i in input_data if i.source != None]


def cleanup_translations(input_data):
    holder = []
    for translation in input_data:
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


def translate_translations(input_data):
    translations = Translator().translate(input_data, dest="english")
    return [translation.text for translation in translations]


def verification_file(input_data, language):
    output_workbook = Workbook()
    output_workbook.title = "Verification Data"
    output_worksheet = output_workbook.active
    for x in range(1, 3):
        for y in range(1, len(input_data) + 1):
            output_worksheet.cell(row=x, column=y)

    output_worksheet["A1"] = "Source"
    output_worksheet["B1"] = "Translated Translation"

    index = 2
    for translation in input_data:
        if translation.source != None:
            output_worksheet["A{}".format(index)] = translation.source
            output_worksheet["B{}".format(index)] = translation.verification_text
            index += 1

    output_workbook.save("Verification_Data_{}.xlsx".format(language))


def move_to_ts(input_data, language):
    print("this is where I'd move the data to the TS file...\n\n\n\n\n\nIF I HAD ONE")


def main():

    # explication
    excel_file, verification_mode = get_args()

    # explication
    data = get_data(excel_file) if excel_file else None
    lang = excel_file[-7:-5]

    # explication
    if data:

        # explication
        if verification_mode:

            # explication
            foreign_words = get_translations(data)

            # explication
            foreign_words = cleanup_translations(foreign_words)

            # explication
            foreign_words = translate_translations(foreign_words)

            # explication
            index = 0
            for translation in data:

                # explication
                if translation.source != None:
                    translation.verification_text = foreign_words[index]
                    index += 1

            # explication
            verification_file(data, lang)

        # explication
        move_to_ts(data, lang)


if __name__ == "__main__":
    main()