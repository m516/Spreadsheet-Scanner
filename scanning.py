'''
A script we can use to make sure all the lab submissions we get are original.
 
Usage:
1. Get Python, at least 3.10
2. Get dependencies with python -m pip install levenshtein openpyxl
3. Download all the submissions. The script can filter out any duplicates with the same values for the name cell.
4. Place this script in that folder. Configure the options on lines 20-50 if need be
5. Run the script. It may take a few minutes.
'''

# External deps
import os                            # OS
from openpyxl import load_workbook   # Reading Excel files
from rapidfuzz.distance import Indel # Normalized Indel similarity between two strings
 


######################################################################################################################
'''
 _____              __ _                       _   _             
/  __ \            / _(_)                     | | (_)            
| /  \/ ___  _ __ | |_ _  __ _ _   _ _ __ __ _| |_ _  ___  _ __  
| |    / _ \| '_ \|  _| |/ _` | | | | '__/ _` | __| |/ _ \| '_ \ 
| \__/\ (_) | | | | | | | (_| | |_| | | | (_| | |_| | (_) | | | |
 \____/\___/|_| |_|_| |_|\__, |\__,_|_|  \__,_|\__|_|\___/|_| |_|
                          __/ |                                  
                         |___/                                   
'''


# The path with all the spreadsheets
path = os.path.join(os.getcwd(), 'example') 

# The cell with some unique identifier, like a name.
# This is a tuple with the sheet number (starting at 1), followed by the cell location
nameCell = (1, "B1")       # 1st sheet, B1

# The similarity threshold, a value from 0 to 1.
# The closer this value is to 1, the more picky the algorithm will be.
# The closer it gets to 0, the more similarities you may find.
similarityThreshold = 0.9  # Filter results to at least 90% similarity

# A list of all the cells that have some answer that should be specific to each name.
# Same structure as the name cell.
cellsToCheck = [
    (1, "A5")              # 1st sheet, A5
]
#######################################################################################################################







# Utils
title = '''                                                                                
      _/_/_/                                          _/                               
   _/          _/_/_/    _/_/_/  _/_/_/    _/_/_/        _/_/_/      _/_/_/            
    _/_/    _/        _/    _/  _/    _/  _/    _/  _/  _/    _/  _/    _/             
       _/  _/        _/    _/  _/    _/  _/    _/  _/  _/    _/  _/    _/              
_/_/_/      _/_/_/    _/_/_/  _/    _/  _/    _/  _/  _/    _/    _/_/_/  _/  _/  _/   
                                                                     _/                
                                                                _/_/                   
                                           Grab some coffee; this could take a while.
'''
version = "1.0"

def color_text(text, r, g, b):
    '''
    Converts text into colored text
    https://stackoverflow.com/questions/287871/how-to-print-colored-text-in-python
    '''
    return "\033[38;2;{};{};{}m{} \033[38;2;255;255;255m".format(r, g, b, text)

def print_intro():
    '''
    Prints an intro message with the ASCII art title and version
    '''
    print(color_text(title, 255, 64, 0))
    print(color_text("Version " + version, 255, 64, 0))

def similarity(string1, string2):
    '''
    Given two strings, returns a value between 0 and 1.
    0 means there is absolutely no similarity between the two strings.
    1 means the strings are identical. 
    '''
    return Indel.normalized_similarity(string1, string2)




# Globals
dataToCheck = []
hits = 0



# Main script
print_intro()
 
# Get all Excel spreadsheets in a folder
for x in os.listdir(path):
    if x.endswith(".xlsx"): # For each spreadsheet
        # Open the workbook
        # wb = xlrd.open_workbook(x) # When the lab modernizes some day
        wb = load_workbook(filename = os.path.join(path, x))

        # Get the names of the students submitting
        name = wb.worksheets[nameCell[0]-1][nameCell[1]].value

        # Get the elements to check in all the spreadsheets
        data = []
        for i in cellsToCheck:
            content = wb.worksheets[i[0]-1][i[1]].value
            data.append({
                "location": i,
                "content" : str(content)
            })

        # Package the data into a spreadsheet
        d = {
            "name": name,
            "data": data
        }

        # Print for debugging
        # print(d)
        print('.', end='', flush=True)

        # Add to the list if not duplicate
        if d["name"] not in (i["name"] for i in dataToCheck):
            dataToCheck.append(d)

print("Done reading spreadsheets. Processing...")

for submission1 in dataToCheck:
    for submission2 in dataToCheck:
        # Skip redundant checks
        if submission1["name"] <= submission2["name"]:
            # Don't check, no copying here
            continue
        
        
        # For each item to check
        for i in range(len(submission1["data"])):
            string1 = submission1["data"][i]["content"]
            string2 = submission2["data"][i]["content"]
            s = similarity(string1, string2)
            if(s > similarityThreshold):
                print(color_text("---------------------------  MATCH DETECTED  --------------------------", 255, 0, 0))
                print('''
Names: 
    %s
    %s
Location:   %s
Similarity: %s
Response from %s:
    %s
Response from %s:
    %s





''' % (
                    color_text(submission1["name"],                  0, 128, 255),
                    color_text(submission2["name"],                255, 128,   0),
                    color_text(submission1["data"][i]["location"], 255,   0, 255),
                    color_text(str(s),                             255,   0, 255),
                    color_text(submission1["name"],                  0, 128, 255),
                    color_text(submission1["data"][i]["content"],    0, 128, 255),
                    color_text(submission2["name"],                255, 128,   0),
                    color_text(submission2["data"][i]["content"],  255, 128,   0),
                    )
                )
                hits += 1

print(
'''
--------------------  Conclusion  -------------------
%s matches found.
Done
''' % (str(hits))
)           