import googletrans
import openpyxl
from googletrans import Translator

# Translate the first column in the excel file into whatever language is needed and put it into the second column

# take in a list of sentences and translate them
def translate(sentences: list, desiredLang: str):
    translator = Translator()
    translations = translator.translate(sentences, dest=desiredLang)
    return [translation.text for translation in translations]


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

    def data(self):
        return (
            self.tagStart,
            self.typ,
            self.index,
            self.contentStart,
            self.contentStart,
            self.content,
        )


fileName = "exspiron2xi"

# get the desired language
deciding = True
while deciding:
    desiredLang = input(
        'what is the desired lang?\nenter "1" if you need to see all the language codes\n'
    )
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
    with open("{}_fr.ts".format(fileName), "r", encoding="utf-8") as inputData:
        lineIndex = 0

        # loop through each line
        for line in inputData:

            # find the specific tags
            if "<source" in line or "<translation" in line:

                # set these up outside class definition for use in class definition
                contentStart = line.find(">") + 1
                contentEnd = line.find("<", contentStart)

                # build the list of tags
                extractedTags.append(
                    tag(
                        line.find("<"),
                        line[line.find("<") + 1 : line.find(">")],
                        lineIndex,
                        contentStart,
                        contentEnd,
                        line[contentStart:contentEnd],
                    )
                )
            lineIndex += 1

    # get the raw data as a list
    data = open("{}_fr.ts".format(fileName), "r", encoding="utf-8").readlines()

    # create an empty file with the correct naming scheme or edit it if it somehow already exists
    try:
        open("{}_{}.ts".format(fileName, desiredLang), "w", encoding="utf-8")
    except:
        print("no output file found\ncreating one")
        open("{}_{}.ts".format(fileName, desiredLang), "x", encoding="utf-8").close()

    # load the output file
    with open(
        "{}_{}.ts".format(fileName, desiredLang), "w", encoding="utf-8"
    ) as outputData:

        # create a list of the things that need translated, the locations of their destinations, and the index of the class in the extracted list
        dct = [
            [source.content for source in extractedTags if source.typ == "source"],
            [
                translation.index[0]
                for translation in extractedTags
                if translation.typ == "translation"
                or translation.typ == 'translation type="unfinished"'
            ],
            [
                i
                for i in range(len(extractedTags))
                if extractedTags[i].typ == "translation"
                or extractedTags[i].typ == 'translation type="unfinished"'
            ],
        ]

        # translate their text
        translations = translate(dct[0], desiredLang)

        # replace the text in the raw data
        for index, location in enumerate(dct[1]):
            if (
                extractedTags[dct[2][index]].contentStart
                != extractedTags[dct[2][index]].contentEnd
            ):
                data[location] = data[location].replace(
                    extractedTags[dct[2][index]].content, translations[index]
                )
            else:
                data[location] = (
                    data[location][: extractedTags[dct[2][index]].contentStart]
                    + str(translations[index])
                    + data[location][extractedTags[dct[2][index]].contentStart :]
                )

        # write to the file
        outputData.writelines(data)


# generate an example file if one is not found where expected
except FileNotFoundError as error:
    print("this file does not exist\nCreating one where it's expected\n")
    open("place the file here named '{}'.ts".format(fileName), "x", encoding="utf-8")
    input("press enter to close\n")
