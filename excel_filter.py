import pandas as pd  # pip install pandas
import os


file_name = "FP.xlsx"  # file name with .xlsx and maybe .xls at the end. Must be in the same folder
# crate sheet2 in excel, create a column named "Remove" and place all missing PIDs in that column.

# Make sure the names below are exactly the same as in the Excel spreadsheet, it's CASE SENSITIVE!

Full_id_column = "new_id"  # column name for the IDs to filter, it's case-sensitive!
Ids_to_be_removed = "remove"  # column name for the IDs to be removed from the column above, it's case-sensitive!
Sheet1_name = 'Sheet1'  # these are case-sensitive also, just paste the sheet
# name with the full list of ids, it's case-sensitive!
Sheet2_name = 'Sheet2'  # paste the sheet name which contains IDs to remove, it's case-sensitive!

# try except part will create 2 folders, new and old.
# Old file will be copied to the old folder and new file will be created in the new folder
# If they already exist, it just prints out that they exist, and it won't do anything.
try:
    directory_new = 'new'  # sets the folder name
    directory_old = 'old'  # sets the folder name
    parent_dir = os.getcwd()  # gets the parent directory
    path = os.path.join(parent_dir, directory_new)  # creates a path to be created like c:/users/user/downloads/new/
    os.mkdir(path)  # creates the "new" folder based on path, so
    # it will be always in the same main folder in which this .py file is
    path = os.path.join(parent_dir, directory_old)  # same as above, it sets the path for folder "old"
    os.mkdir(path)  # creates folder "old" in the same directory as this file

except FileExistsError:
    print("directory 'new' and 'old' already present")  # prints out a message if directories were created previously
    pass  # ends the exception and continues running the program


def remove_dupes():  # main function is defined.
    pd.set_option('display.max_rows', 10000)  # increase the int value if there are more than 10k rows.
    data = pd.read_excel(file_name, sheet_name=Sheet1_name)  # reads the Excel files 1st sheet
    data2 = pd.read_excel(file_name, sheet_name=Sheet2_name)  # reads the Excel files 2nd sheet
    condition = data[Full_id_column].str.upper().isin(data2[Ids_to_be_removed].str.upper())  # sets the condition, if ids from 2nd sheet
    # are in the first sheet, it marks the indexes
    # print(condition)
    # print(data[condition])
    new = data.drop(data.index[condition], axis=0)  # deletes Excel rows that were marked by the condition
    # print(new)
    dataframe_final = new  # just assigns it to a new variable, fewer changes when working on this program
    # print(dataframe_final)
    dataframe_final.to_excel(os.getcwd() + "/new/new_" + file_name)  # exports the dataframe to an
    # Excel file and places it in the new folder under the name Missing_PIDs_removed.xlsx


remove_dupes()  # function is called to run the program
# uncomment the line below by deleting the 'hashtag' and whitespace to move the original file to the old folder
# os.rename(os.getcwd()+"/"+file_name, os.getcwd()+"/old/"+file_name)
