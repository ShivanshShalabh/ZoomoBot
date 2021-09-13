import re
import os
import sys


def isValidHexaCode(str):
    # Regex to check valid
    # hexadecimal color code.
    regex = "^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$"

    # Compile the ReGex
    p = re.compile(regex)

    # If the string is empty
    # return false
    if(str == None):
        return False

    # Return if the string
    # matched the ReGex
    if(re.search(p, str)):
        return True
    else:
        return False


if __name__ == '__main__':
    # Check if os is windows
    if os.name == 'nt':
        excecutable_file = 'chromedriver.exe'
    else:
        excecutable_file = 'chromedriver'
    # Check if file with name Cache.txt exists
    if not os.path.isfile(excecutable_file):
        print("Error: Make sure ChromeDriver is downloaded and the",
              excecutable_file, "is in the folder")
        print("To download the file, visit https://chromedriver.chromium.org/downloads")
        sys.exit()
    if os.path.isfile('Cache.txt'):
        # If file exists, delete it
        os.remove('Cache.txt')

    # Create file with name Cache.txt
    name = input('Enter your name (Enter -1 to skip): ')
    while not name:
        name = input('Enter your name (Enter -1 to skip): ')
    if name == '-1':
        name = ''
    skip_column = 0
    skip_row = 0
    print("Make sure that the names in the Excel file are in the same column without any empty cell in between.")
    cell = input('Enter cell address to start reading from: ')
    while not cell:
        cell = input('Enter cell address to start reading from: ')
    for i in cell:
        if i.isalpha():

            skip_column += ord(i.upper())-65

        elif i.isdigit():
            skip_row += int(i)-1

    # Write name, skip_column and skip_row to Cache.txt
    color = input(
        'Enter hex value of the color with which you want to color the cell (Enter -1 to skip): ')
    while not isValidHexaCode(color) and (color != '-1'):
        color = input(
            'Enter hex value of the color with which you want to color the cell (Enter -1 to skip): ')
    if color == '-1':
        color = ''
    skip_column, skip_row = str(skip_column), str(skip_row)
    with open('Cache.txt', 'w') as f:
        f.write(name + "|Name" + '\n' + skip_column +
                "|No. of columns to skip" + '\n' + skip_row + '|No. of rows to skip\n'+color+'|Cell Color')
