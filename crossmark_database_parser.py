# INTEL CONFIDENTIAL
# Copyright 2015 Intel Corporation All Rights Reserved.
#
# The source code contained or described herein and all documents related
# to the source code ("Material") are owned by Intel Corporation or its
# suppliers or licensors. Title to the Material remains with Intel Corp-
# oration or its suppliers and licensors. The Material may contain trade
# secrets and proprietary and confidential information of Intel Corpor-
# ation and its suppliers and licensors, and is protected by worldwide
# copyright and trade secret laws and treaty provisions. No part of the
# Material may be used, copied, reproduced, modified, published, uploaded,
# posted, transmitted, distributed, or disclosed in any way without
# Intel's prior express written permission.
#
# No license under any patent, copyright, trade secret or other intellect-
# ual property right is granted to or conferred upon you by disclosure or
# delivery of the Materials, either expressly, by implication, inducement,
# estoppel or otherwise. Any license under such intellectual property
# rights must be express and approved by Intel in writing.

# Author: Ronny Z. Valtonen
# Date Created: 07/17/2023
# Purpose: To assist on parsing files created by Crossmark Debug Tool.

#####################################
# BASIC PREP                        #
# pip install openpyxl (3.0.10)     #
# pip install xlsxwriter (3.0.3)    #
# pip install pandas (1.4.2)        #
# python 3.9.12                     #
# Linux: sudo pacman -S tk          #
# MacOS: brew install python-tk     #
#####################################

# Program
import os
import xlsxwriter
import csv

# UI
from tkinter import filedialog
from tkinter import *

# Parses the Crossmark subtest scores.
def parse_subtests(file, sheet):
    pass

# Parses the initial Crossmark test scores.
def parse_performance(file, sheet):
    print("Found initial Crossmark performance csv, beginning data collection.")
    # Collect the data here
    # Create a worksheet to write to.
    my_worksheet = sheet.add_worksheet()

    # with open(file, newline = '') as csvfile:
    #     main_reader = csv.reader(csvfile, delimiter = ' ', quotechar = '|')
    #     for row in main_reader:
    #         print(', '.join(row))

    with open(file) as main_scores:
        reader = csv.reader(main_scores)
        rows = list(reader)
        # Printing main scores
        print("Printing Main Scores...")
        print(rows)
    
    # Get scores
    real_row_1 = rows[6]
    real_row_2 = rows[7]
    real_row_3 = rows[8]
    real_row_4 = rows[9]
    crossmark_score = real_row_1[4]
    productivity_score = real_row_2[4]
    creativity_score = real_row_3[4]
    responsiveness_score = real_row_4[4]

    # Get run folder
    split_data = file.split("\\")
    my_file = split_data[-2]

    # Once done, parse the subtest csv file.
    print("Grabbing subtests")

    # Convert the file into a string, remove the last scores.csv characters, and add in the desired csv file we want next.
    convert_to_string = str(file)
    previous_directory = (convert_to_string[0:-10])
    previous_directory += "measure_performance.csv"
    print(previous_directory)

    with open(previous_directory) as sub_scores:
        reader_second = csv.reader(sub_scores)
        sub_rows = list(reader_second)
        # Printed sub scores
        print("Printing Sub-Test Scores")
        print(sub_rows)

    # Write to the csv now.
    my_worksheet.write('A1', my_file)
    my_worksheet.write('B1', 'Crossmark Overall Score')
    my_worksheet.write('C1', crossmark_score)
    my_worksheet.write('A2', my_file)
    my_worksheet.write('B2', 'Productivity Score')
    my_worksheet.write('C2', productivity_score)
    my_worksheet.write('A3', my_file)
    my_worksheet.write('B3', 'Creativity Score')
    my_worksheet.write('C3', creativity_score)
    my_worksheet.write('A4', my_file)
    my_worksheet.write('B4', 'Responsiveness Score')
    my_worksheet.write('C4', responsiveness_score)

# Selects the appropriate files within directory.
def pick_file(window, selected_file, workbook):
    if "scores.csv" in selected_file:
        parse_performance(selected_file, workbook)


# Program driver.
def main():
    pass
    # Declare a window
    window = Tk()
    print("Prompting user with File Explorer")

    window.filename = filedialog.askdirectory(initialdir= "/", title = "Select Debug Folder")
    window.geometry("750x270")

    # Initialize label
    Label(window, text = "Parsing Data, Please Wait... \n Created by Ronny V.", font=('Helvetica 20 bold')).pack(pady=20)
    window.after(1000, lambda: window.destroy())
    window.mainloop()

    if os.path.exists('debug.xlsx'):
        os.remove('debug.xlsx')

    if os.path.exists('database.xlsx'):
        os.remove('database.xlsx')

    workbook = xlsxwriter.Workbook('debug.xlsx')

    for root, dirs, files in os.walk(window.filename):
        for file in files:
            try:
                selected_file = os.path.join(root, file)
                pick_file(window, selected_file, workbook)

            except:
                pass

    workbook.close()

# Runs program in correct order.
if __name__ == "__main__":
    print("Starting program")
    main()
