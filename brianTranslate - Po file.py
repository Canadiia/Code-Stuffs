import googletrans
import openpyxl
from googletrans import Translator

"""This will take in a *clean* .po file and translate it into a language specified

#: ../example/directory/here
msgctxt "example"

#! The MSGID needs to be on one line only
msgid "Access Service Mode"

#* will output as one line
msgstr "Accéder au mode de service"
"""


def translate(sentences: list, desiredLang: str):
    """takes a list of strings and translates them into the desired language

    Args:
        sentences (list): the list of sentences needing translated
        desiredLang (str): the language you want to translate into
            !! MUST BE IN THE googletrans.LANGCODES LIST !!

    Returns:
        Translated sentencen: the list of sentences after they've been translated
    """
    translator = Translator()
    translations = translator.translate(sentences, dest=desiredLang)
    return [translation.text for translation in translations]


class tag:
    """
    holds some useful information about a tag
    """

    def __init__(
        self,
        typ: str,
        index: int,
    ):
        self.tagStart = 0
        self.typ = typ
        self.index = [index, [], []]
        self.content = None


#! this is where the name of the blank file goes
fileName = "exspiron2xi"

# get the desired language
deciding = True
while deciding:
    desiredLang = input(
        'what is the desired lang?\nenter "1" if you need to see all the language codes\n'
    )

    # check if what was inputted is actually a valid language code
    # or display all the language codes for reference
    if desiredLang.isdigit():
        if int(desiredLang) == 1:
            for lang in googletrans.LANGCODES:
                print(lang)
        else:
            print("Invalid Input")
    else:
        desiredLang = desiredLang.lower()
        if desiredLang in googletrans.LANGCODES:
            deciding = False
        else:
            print("invalid language code")

try:
    extractedTags = []

    # load the data
    # ? this takes in the blank file, and uses it exclusively for reading
    with open("{}_fr.po".format(fileName), "r", encoding="utf-8") as inputData:
        lineIndex = 0

        # loop through each line
        for index, line in enumerate(inputData):

            # find the specific tags and append a tag class with what type of tag it is, and where it is in the raw data list
            if "msgid" in line:
                extractedTags.append(tag("msgid", lineIndex))
            elif "msgstr" in line:
                extractedTags.append(tag("msgstr", lineIndex))

            # capture blank space lines for use in splitting the data further down
            elif line == "\n":
                extractedTags.append(tag("line break", lineIndex))
            lineIndex += 1

        # delete the first 3 tags as they are the two blank "translations" at the very top of the page along with the first blank line
        del extractedTags[0:3]

        # get the raw data as a list
        data = open("{}_fr.po".format(fileName), "r", encoding="utf-8").readlines()

        # loop through all the tags with indices for indexing later
        for index, tg in enumerate(extractedTags):

            # get what the next tag is for reference
            nextTag = extractedTags[
                index + 1 if not index + 1 >= len(extractedTags) else index
            ]

            # exclude line break tags
            if tg.typ != "line break":

                # set the content of the tag to be everything between the current tag and the next tag (including line breaks)
                tg.content = [
                    sentence for sentence in data[tg.index[0] : nextTag.index[0]]
                ]

                # set the index value of the tag to the location of the start quotation mark and ending quotation mark
                # this will allow us to place the translation into where the quote is
                tg.index[1] = [str(string).find('"') for string in tg.content]
                tg.index[2] = [
                    str(string).find('"', tg.index[1][-1] + 1) for string in tg.content
                ]

        # we're done with all the "line break" tags so remove them
        for index in range(len(extractedTags)):
            if not index >= len(extractedTags):
                if extractedTags[index].typ == "line break":
                    del extractedTags[index]

    # create an empty output file with the correct naming scheme or edit it if it somehow already exists
    try:
        open("{}_{}.po".format(fileName, desiredLang), "w", encoding="utf-8")
    except:
        print("no output file found\ncreating one")
        open("{}_{}.po".format(fileName, desiredLang), "x", encoding="utf-8").close()

    # load the output file
    with open(
        "{}_{}.po".format(fileName, desiredLang), "w", encoding="utf-8"
    ) as outputData:

        # create a list of the things that need translated (empty at first),
        # the locations of their destinations,
        # and the index of the class in the extracted list
        dct = [
            [],
            [
                translation.index[0]
                for translation in extractedTags
                if translation.typ == "msgstr"
            ],
            [
                index
                for index in range(len(extractedTags))
                if extractedTags[index].typ == "msgstr"
            ],
        ]

        # actually put the data needing translated into the dct list
        for string in extractedTags:
            if string.typ == "msgid":
                dct[0].append(
                    [
                        string.content[index][
                            string.index[1][index] + 1 : string.index[2][index]
                        ]
                        for index in range(len(string.content))
                    ]
                )

        # translate their text
        for index, sentences in enumerate(dct[0]):
            dct[0][index] = translate(sentences, desiredLang)

        # fix the broken line breaks
        for index, sentence in enumerate(dct[0]):
            prevTag = extractedTags[dct[2][index] - 1]
            for jndex, lne in enumerate(sentence):
                if "\\ n" in lne.lower():
                    dct[0][index][jndex] = lne.replace("\\ n", "\\n").replace(
                        "\\ N", "\\n"
                    )
                if '\\"%1\\"' in prevTag.content[0]:
                    dct[0][index][jndex] = lne.replace(" \\ ", '\\"%1\\"', 1)

        # replace the text in the raw data
        # gonna need this range somehow
        # *  data[tg.index[0] : nextTag.index[0]]
        # *  nextTag = extractedTags[
        # *     index + 1 if not index + 1 >= len(extractedTags) else index
        # *  ]
        for index, location in enumerate(dct[1]):
            if (
                extractedTags[dct[2][index]].index[1][0] + 1
                != extractedTags[dct[2][index]].index[2][0]
            ):
                data[location] = data[location].replace(
                    extractedTags[dct[2][index]].content[0][
                        extractedTags[dct[2][index]].index[1][0]
                        + 1 : extractedTags[dct[2][index]].index[2][0]
                    ],
                    dct[0][index][0],
                )
            elif index < len(data):
                data[location] = (
                    data[location][: extractedTags[dct[2][index]].index[1][0] + 1]
                    + str(dct[0][index][0])
                    + data[location][extractedTags[dct[2][index]].index[1][0] + 1 :]
                )

        # write to the file
        outputData.writelines(data)


# generate an example file if one is not found where expected
except FileNotFoundError as error:
    print("this file does not exist\nCreating one where it's expected\n")
    open("place the file here named '{}'.po".format(fileName), "x", encoding="utf-8")
    input("press enter to close\n")
