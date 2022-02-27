# Imports all required modules
# If when run there is an error that says "module not found"
# go to command prompt/terminal and type:
# pip install <module>
# If the module missing is win32com, pip install pywin32
from libcpp cimport bool
import tabula
import pandas as pd
import win32com.client as win32
import os
import numpy as np
import sys
cimport cython
import time
import getpass
import shutil


cdef str billing_book, path, file, row, current_dir, specific_activity, RANGE, tab, price, price2, price3, j, current_price
cdef int color
cdef double price_one, price_two, price_three, current_Price
cdef list frames, files, inds, name2, prices
cdef bool go


# Delete gen_py folder as can be corrupt and fail script
username = getpass.getuser()
try:
    gen_py = f'C:\\Users\\{username}\\AppData\\Local\\Temp\\gen_py'
    shutil.rmtree(gen_py)
except FileNotFoundError:
    pass

# The following finds the path of the current directory
# This is used as the base of where to read files and where to save new files  created
current_dir = os.getcwd()
files = os.listdir()
for file in files:
    if ".pdf" in file:
        path = file
    if "consolidated" in file.lower():
        billing_book = current_dir + "\\" + file

# For the user to tell the program how to fill in the main spreadsheet
# Comment out the 2 input statement and un-comment the 2 lines below if you can't be bothered to re-input the same name& po_number over and over is all POs have the same details
# Remember to change back after use
po_number = input("Please Enter the full PO reference to be used: ")
contact_name = input("Please enter contact name: ")
pdf_pages = input("How many pages in PDF: ")

if pdf_pages == "":
    pdf_pages = 10

# If a specific activity is indicated on the PO, let the program know to save wasted row reading
# Leave blank if unknown or mixed
specific_activity = input("""\nIs it a specific activity?\n
Internal Cleaning: ic
External Cleaning: ex
Cleaning: cl
Grounds Maintenance: gm
Window Cleaning: wc
Concierge Services: cs
No Specific Activity: (leave blank)\n
""")

specific_lot = input("""\nIs it a specific lot?\n
Lot A: a
Lot B: b
Lot C: c
Lot D: d
NHG: nhg
No Specific Lot: (leave blank)\n
""")

# Creates the dataframe used by reading just one page from the pdf
df = tabula.read_pdf(path, pages='all', lattice=True)[1]


# Updates the main dataframe but creating a new dataframe for each pdf page and concatening the 2 at the end of each round
# Will attempt to read 20 pages
# If there are less than 20, the funtion will just say "Index out of range" and continue
# If there are more, update range(20) to range(25..30..35 etc..)
# Make sure the number is at least 2 higher than you think you need

cdef int page
cdef int pdf_page
page = 0
for pdf_page in range(int(pdf_pages)+2):
    try:
        print(f"Reading pdf page {page}")
        df2 = tabula.read_pdf(path, pages='all', lattice=True)[page]
        frames = [df, df2]
        if page == 1:
            df = df2
        else:
            df = pd.concat(frames)

    except IndexError:
        print("Index out of range")

    page += 1


# Changes the index to a normal 1-n as it is currently 1,2,3,1,2,3,1,2,3 due to the concatenation of multiple indiviual dataframes
df = df.reset_index()
del df['index']


path = path.replace(".pdf", "")
# Saves the dataframe as a new file to be read later
df.to_excel(current_dir + f'\\{path} nhhg.xlsx')
# Reads the dataframe just saved
nhg = pd.read_excel(current_dir + f'\\{path} nhhg.xlsx')


# One cell in the nhhg dataframe turns to a specific color to indicate but the row has been completed successfully
# and which tab/lot it was found in
# Orange
# color = 111327
# Strong Blue
# color = 16711680
# Cyan
# color = 16776960
# Light Green
color = 9359529

# Indicated which columns to read and/or enter data into
# int_net = monthly net price
# int_po = po_number
# int_name = name of provider (e.g. Caroline, Stuart)
int_net = "AL"
int_po = "AP"
int_name = "AQ"


# c internal cleaning = 1-1104
# b internal cleaning = 1105-1333
# a internal cleaning = 1334-1479
# nhg internal cleaning = 1480-2003
# d internal cleaning = 2004-2126

# internal cleaning
# external cleaning
# window cleaning
# grounds maintenance
# concierge services

# c = 1104
# b = 229
# a = 146
# nhg = 524
# d = 123


# function returns a tuple
# first element is first_row number
# second element is last_row number
def specific_lot_range(specific_lot="", activity_starting_row=1, activity_ending_row=11000):
    if specific_lot == "c":
        first_row = activity_starting_row
        last_row = activity_starting_row + 1104
    elif specific_lot == "b":
        first_row = activity_starting_row + 1104
        last_row = activity_starting_row + 1104 + 229
    elif specific_lot == "a":
        first_row = activity_starting_row + 1104 + 229
        last_row = activity_starting_row + 1104 + 229 + 146
    elif specific_lot == "nhg":
        first_row = activity_starting_row + 1104 + 229 + 146
        last_row = activity_starting_row + 1104 + 229 + 146 + 524
    elif specific_lot == "d":
        first_row = activity_ending_row - 123
        last_row = activity_ending_row
    else:
        first_row = activity_starting_row
        last_row = activity_ending_row

    return first_row, last_row


# Provides deafualt range
# Can probably be removed
if specific_activity == "ic":
    RANGE = "G1:H2126"
    a = specific_lot_range(specific_lot, 1, 2126)
    RANGE = f"G{a[0]}:H{a[1]}"

elif specific_activity == "ec":
    RANGE = "G2127:H4248"
    a = specific_lot_range(specific_lot, 2127, 4248)
    RANGE = f"G{a[0]}:H{a[1]}"

elif specific_activity == "cl":
    RANGE = "G1:H4248"
    a = specific_lot_range(specific_lot, 1, 4248)
    RANGE = f"G{a[0]}:H{a[1]}"

elif specific_activity == "wc":
    RANGE = "G4249:H6363"
    a = specific_lot_range(specific_lot, 4249, 6363)
    RANGE = f"G{a[0]}:H{a[1]}"

elif specific_activity == "gm":
    RANGE = "G6364:H8479"
    a = specific_lot_range(specific_lot, 6364, 8479)
    RANGE = f"G{a[0]}:H{a[1]}"

elif specific_activity == "cs":
    RANGE = "G8480:H10700"
    a = specific_lot_range(specific_lot, 8480, 10700)
    RANGE = f"G{a[0]}:H{a[1]}"

else:
    if specific_lot == "":
        RANGE = "G1:H10700"
    else:
        RANGE = "G1:H10700"

print(RANGE)


# Opens the main spreadsheet
xl = win32.gencache.EnsureDispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
book = xl.Workbooks.Open(billing_book)

# Gives a second for first workbook to fully open
# might crash if 2nd workbook is opened too quickly
time.sleep(1)

# Reads the dataframe with pdf data on
book2 = xl.Workbooks.Open(current_dir + f'\\{path} nhhg.xlsx')
sheet2 = book2.Sheets('Sheet1')
time.sleep(1)

cdef int row_number

cdef void main():
    row_number = 1
    
    for i in range(len(nhg['Item Code'])):

        # If the correct tab is found, the item is no longer continued as it will waste time
        # go is turned to False so the iteration stops
        go = True

        tab = "OneList"
        c_amended = book.Sheets(tab)

        # Ensure all calculations are made in the spreadsheet so the most recent/correct costs are used
        # c_amended.Calculate()


        # Ignores certain words/grammatic tokens to make comparisons less strict
        name = str(nhg['Description'].iloc[i]).lower().replace(",", " ").replace("london", "").replace(
            "_x000d_", " ").replace("_x000D_", " ").replace("sec06", " ").replace("estate", "")

        if go == True:

            # Turns the address in pdf dataframe into individual words rather than one long address
            name2 = []

            for m in name.split(" "):
                name2.append(m)
            name = list(filter(None, name2))
                               
            # The price on the PO                  
            price_one = np.round(
                float(str(nhg['Cost (£)'].iloc[i]).replace(",", "")), 2)

            # Below comment halves the cost on pdf dataframe in case the cost is over 2 months
            # Uncomment if above comment is true
            # price = np.round(float(str(nhg['Cost (£)'].iloc[i]).replace(",", "")) / 2, 2)
            # Adds give or take £0.01 in case rounds causes an error
            price_two = np.round(price_one + 0.01, 2)
            price_three = np.round(price_one - 0.01, 2)
            price = str(price_one)
            price2 = str(price_two)
            price3 = str(price_three)
            prices = [price, price2, price3]

            print(f"Looking for: {name}, Price : {price}")

            if name[0] == "nan" and name[-1] == "nan":
                sys.exit()

            # Turns the address in main spreadsheet into individual words rather than one long address
            for q in c_amended.Range(RANGE):
                sheet_name = str(q.Value).lower().replace(",", " ").replace(
                    "london", '').replace("_x000d_", " ").replace("_x000D_", " ")
                split_name = []


                for j in sheet_name.split(" "):
                    split_name.append(j)
                split_name = list(filter(None, split_name))

                if all(x in split_name for x in name) and go == True:
                    row = str(q.Row)
                    
                    
                    try:
                        current_Price = float(
                            c_amended.Range(f"{int_net}{row}").Value)
                    except:
                        current_Price = 0.

                    current_Price = np.round(current_Price, 2)
                    current_price = str(current_Price)
                    print(f"Row {row}: £{current_price}")

                    # print(current_price)
                    # The following is where prices are compared between pdf dataframe and main spreadsheet to find the correct column/activity
                    try:
                        if go == True and current_price in prices and "" in str(c_amended.Range(f"B{row}").Value).lower():

                            # try:
                            c_amended.Range(
                                f"{int_po}{row}").Value = po_number
                            # except Exception as e:
                            #     print(int_po, row, po_number)
                            #     print(e)
                            c_amended.Range(
                                f"{int_name}{row}").Value = contact_name

                            print(f"Found {split_name}, price {current_price}")
                            sheet2.Cells(i+2, 2).Interior.Color = color
                            go = False
                            break

                    except TypeError as e:
                        print(e)


        row_number += 1

        if row_number % 10 == 0:
            # Saves both dataframes every 10 iterations in case the program crashes during the middle/near the end of a long/patient run!!!
            # Urgghh, would be very annoying if not done!!
            book.Save()
            book2.Save()

main()

print("Completed")

# Thank you for reading, and your welcome Courtney
# God, I'm too old for this crap...
