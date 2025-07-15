import pandas as pd
import re
import os
from openpyxl import Workbook

path_of_ExeFile = os.path.abspath(__file__)  # Absolute path to executable
path_of_ExeFile = path_of_ExeFile.replace("MainSplit.py", "text_files\\")  # Changing path where's .txt files located


def find_txt(dir):  # The function can get all txt files in the directory "text_files"
    files = []
    for file in os.listdir(dir):
        if file.endswith(".txt"):
            files.append(os.path.join(file))
    return files


print(path_of_ExeFile)
Text_Files = find_txt(os.path.dirname(path_of_ExeFile))  # getting all txt files from directory

Youtube_Regex_Pattern = r'https?://(?:www\.)?(?:youtube\.com/watch\?v=|youtu\.be/)([a-zA-Z0-9_-]+)'  # Finding all YouTube ID's from links


def Find_Links(pattern, text):
    matches = re.findall(pattern, text)
    return matches


for file in Text_Files:
    S = open(f'text_files/{file}', encoding="utf8").read()
    IDs = Find_Links(Youtube_Regex_Pattern, S)
    Links = []
    for ID in IDs:
        Links.append(f'https://youtu.be/{ID}')
    df = pd.DataFrame(
        {"Links": [link for link in Links]}
    )

    df.to_excel(f'table_{file}.xlsx', sheet_name="Links", index=False)