import pandas as pd
import re
import os
import sys
from pathlib import Path
from openpyxl import Workbook

if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
if getattr(sys, "frozen", False):
        path_of_ExeFile = str(Path(sys.executable).parent)
else:path_of_ExeFile = str(Path(__file__).resolve().parent) # Absolute path to executable

print("Throw your .txt files in folder where Executable is located")
if not path_of_ExeFile.endswith(os.sep):
    path_of_ExeFile += os.sep
print(path_of_ExeFile)
input("Press enter if you're ready")

if sys.platform == "win32":
    import ctypes

    ctypes.windll.kernel32.SetDllDirectoryW(None)


def find_txt(dir):  # The function can get all txt files
    files = []
    for file in os.listdir(dir):
        if str(file).endswith(".txt"):
            print(file)
            files.append(os.path.join(file))
    return files


print(path_of_ExeFile)
Text_Files = find_txt(os.path.dirname(path_of_ExeFile))  # getting all txt files from directory
Youtube_Regex_Pattern = r'https?://(?:www\.)?(?:youtube\.com/watch\?v=|youtu\.be/)([a-zA-Z0-9_-]+)'  # Finding all YouTube ID's from links


def Find_Links(pattern, text):
    matches = re.findall(pattern, text)
    return matches


for file in Text_Files:
    S = open(f'{file}', encoding="utf8").read()
    IDs = Find_Links(Youtube_Regex_Pattern, S)
    Links = []
    for ID in IDs:
        Links.append(f'https://youtu.be/{ID}')
    df = pd.DataFrame(
        {"Links": [link for link in Links]}
    )

    df.to_excel(f'table_{file[:-4]}.xlsx', sheet_name="Links", index=False)
