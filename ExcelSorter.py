import pandas as pd
import numpy as np
import openpyxl as op

search_term=[]
cells_data=[]
indexes = []
final_i = []
inds = []

new_frame = pd.DataFrame()

"""
sheetSelect allow for the user to select which sheet to scan, filepath must be the path of the excel sheet to scan 
"""

def sheetSelect(filepath):
    sheet = str(input("Which sheet would you like to scan?"))
    df = pd.read_excel(filepath, sheet_name=sheet)
    da = pd.DataFrame(df)
    global new_da
    new_da= df.dropna()
    global coll 
    coll = str(input("Which collumn which you like to search in?"))
    search_list = list(new_da[coll])

"""
Search_terms will create a list that contains terms to search for 
"""

def search_terms():
    while True:
        search_term=[]
        print("When finished entering search terms, please type 'done'.")
        x = str(input("Enter search term:"))
        search_term.append(x)
        if x == "done":
            search_term.remove("done")
            break
    print(search_term)

"""
Create_list will compare terms from the list then create a new Pandas frame for it.
Then it creates an entirely next excel file. 
"""

def create_list():
    for cell in new_da[coll]:
        for term in search_term:
            if term in cell:
                cells_data.append(cell)
    print(cells_data)

"""
Determine_index will create a list of indexes containg the terms
"""

def determine_index():
    for terms in cells_data:
        indexes.append(new_da.index[new_da[coll] == terms].to_list())
    for lists in indexes:
        for nums in lists:
            final_i.append(nums)
    for indx in final_i:
        if indx not in inds:
            inds.append(indx)
"""

Creates a new dataframe and excel file
"""

def dataframe_create(filename):
    df1 = new_da.loc[inds]
    df1.to_excel(filename)

def main(filepath, filename):
    sheetSelect(filepath)
    search_terms()
    create_list()
    determine_index()  
    dataframe_create(filename)         

main()