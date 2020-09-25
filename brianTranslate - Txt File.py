"""
{
{ALARMS_LIMITS_CHANGED_EVENT, "Alarms Limits Changed"},
{ALARMS_PAUSED_EVENT, "Pause signals"},
{SESSION_EVENT, "The session has begun"},
{POOR_SIGNAL_EVENT, "weak signal"},
{PADSET_DISCONNECTED_EVENT, "Sensor Disabled"},
{LOW_BATTERY_EVENT, "Battery Discharged"},
{METRICS_OUT_OF_RANGE_EVENT, "Out Of Range Metrics"},
{POOR_BASELINE_EVENT, "Poor Baseline" },
{EXSPIRON_DISCONNECTED_EVENT, "ExSpiron Disabled"}
}
"""

# ^ The data ^

# Translate the second string into whatever language is needed

import os

import googletrans
from googletrans import Translator
import openpyxl


# get the data needed into a format that is easier to process
def cleanUp(data: list):
    holder = []
    holder2 = []

    # coarse clean up, splitting the two parts
    for line in data:
        holder.append(line.rstrip("\n").split(", "))
    # remove the start and end brackets
    holder.pop(0)
    holder.pop(-1)

    # split the first and second parts into separate lists
    for index, entry in enumerate(holder):
        holder[index] = entry[1]
        holder2.append(entry[0])

    # put the two sorted lists together in the format: [[DATA TO BE TRANSLATED], [THE OTHER STUFF]]
    formattedData = [holder, holder2]

    # pressure wash off all the extra characters
    for index, string in enumerate(formattedData[0]):
        string = string.split('"')
        string = string[1]
        formattedData[0][index] = string
    for index, string in enumerate(formattedData[1]):
        formattedData[1][index] = string.strip("{")

    return formattedData


# take in a list of sentences and translate them
def translate(sentences: list, desiredLang):
    translator = Translator()
    translations = translator.translate(sentences, dest=desiredLang)
    return [translation.text for translation in translations]


# do the final formatting
def reformat(data: list):
    finalData = ["{\n", "}\n"]
    for index in range(len(data[0])):
        combinedString = '{%s, "%s"},\n' % (data[1][index], data[0][index])
        finalData.insert(index + 1, combinedString)
    return finalData


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

# load the data
try:
    inputData = open("Input.txt", "r", encoding="utf-8")
except:
    open("Input.txt", "x", encoding="utf-8").close()
    print(
        "Can't locate input file, creating one where this program expects it to be\nPress enter to close\n"
    )
    input()

# clean up the data
formattedData = cleanUp(inputData)
inputData.close()

# translate the data
formattedData[0] = translate(formattedData[0], desiredLang)

# put the data into a txt file
if not os.path.exists("Output.txt"):
    open("Output.txt", "x", encoding="utf-8").close()
outputData = open("Output.txt", "w", encoding="utf-8")
outputData.writelines(reformat(formattedData))
outputData.close()
input()
