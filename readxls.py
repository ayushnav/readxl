import pandas as pd
import re
import math
import os


def readUserId(df, fileName):
    columnNamesList = df.columns.tolist()
    try:
        pattern = re.compile(r'.*user\s*id.*', re.IGNORECASE)
        columnName = ""
        for key in columnNamesList:
            if pattern.search(key):
                print(key)
                columnName = key
                break

        if(columnName==""):
            print ("USER ID not found in ", fileName)
            return
        useridList = df[key].values
        useridList = [x for x in useridList if type(x) is type("str")]

        print(useridList)

    except Exception as err:
        print(err)


def readDSId(df, fileName):
    columnNamesList = df.columns.tolist()
    try:
        pattern = re.compile(r'.*ds\s*id.*', re.IGNORECASE)
        columnName = ""
        for key in columnNamesList:
            if pattern.search(key):
                print(key)
                columnName = key
                break

        if(columnName==""):
            print ("DS ID not found in ", fileName)
            return
        dsidList = df[key].values
        dsidList = [x for x in dsidList if type(x) is type("str")]
        print(dsidList)

    except Exception as err:
        print(err)

def readEmail(df, fileName):
    columnNamesList = df.columns.tolist()
    emailidList = []
    try:
        pattern = re.compile(r'.*email\s*id.*', re.IGNORECASE)
        columnName = ""

        for key in columnNamesList:
            if pattern.search(key):
                print(key)
                columnName = key
                break
        if columnName=="":
            for col_name in df.columns:
                col_data = df[col_name]

                for data in col_data:
                    if(type(data) == type("str")):
                        # print(data)
                        if "@" in data:
                            emailidList.append(data)
                            print(emailidList)
            return


        if(columnName==""):
            print ("Email ID not found in ", fileName)
            return
        emailidList = df[key].values
        emailidList = [x for x in emailidList if type(x) is type("str")]
        print(emailidList)

    except Exception as err:
        print(err)

def readxls(filename):

    df = pd.read_excel(filename)

    columnNamesList = df.columns.tolist()

    readUserId(df, filename)
    readDSId(df, filename)
    readEmail(df, filename)
    print("Columns List: ", df.columns.tolist())
    print()
    print()
    # print(df.head())

def traverseDir(dirPath):

# List to hold all Excel data frames
    all_data_frames = []

    # Loop through all files in the directory
    for fileName in os.listdir(dirPath):
        # Check if file is an Excel file
        if fileName.startswith("~$"):
            continue
        if fileName.endswith('.xlsx') or fileName.endswith('.xls'):
            # Read Excel file and append data frame to list
            filePath = os.path.join(dirPath, fileName)
            print(fileName)
            readxls(filePath)


if __name__ == '__main__':
    traverseDir(os.getcwd())
