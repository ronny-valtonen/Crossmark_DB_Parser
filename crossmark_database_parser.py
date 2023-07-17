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
        # print(rows)
    
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
        # print(sub_rows)

    # Get scores
    sub_row_1 = sub_rows[23]
    sub_row_2 = sub_rows[24]
    sub_row_3 = sub_rows[25]
    sub_row_4 = sub_rows[26]
    sub_row_5 = sub_rows[27]
    sub_row_6 = sub_rows[28]
    sub_row_7 = sub_rows[29]
    sub_row_8 = sub_rows[30]
    sub_row_9 = sub_rows[31]
    sub_row_10 = sub_rows[32]
    sub_row_11 = sub_rows[33]
    sub_row_12 = sub_rows[34]
    sub_row_13 = sub_rows[35]
    sub_row_14 = sub_rows[36]
    sub_row_15 = sub_rows[37]
    sub_row_16 = sub_rows[38]
    sub_row_17 = sub_rows[39]
    sub_row_18 = sub_rows[40]
    sub_row_19 = sub_rows[41]
    sub_row_20 = sub_rows[42]
    sub_row_21 = sub_rows[43]
    sub_row_22 = sub_rows[44]
    print(sub_rows[23])
    zstd_uncompress_legacy = sub_row_1[-2]
    random_read = sub_row_2[-2]
    black_scholes_serial= sub_row_3[-2]
    string_search = sub_row_4[-2]
    random_write = sub_row_5[-2]
    object_detection = sub_row_6[-2]
    ef_face_recognition = sub_row_7[-2]
    zstd_compress_legacy = sub_row_8[-2]
    zstd_uncompress_streaming = sub_row_9[-2]
    fdt_by_medianflow_tracker = sub_row_10[-2]
    black_scholes_parallel = sub_row_11[-2]
    create_sqlite_blob = sub_row_12[-2]
    video_colorization = sub_row_13[-2]
    external_sort = sub_row_14[-2]
    chacha20_encrypt_openssl = sub_row_15[-2]
    colorization = sub_row_16[-2]
    hdr_stitch = sub_row_17[-2]
    aes_gcm_encrypt_mt = sub_row_18[-2]
    memory_workload = sub_row_19[-2]
    chacha20_decrypt_openssl = sub_row_20[-2]
    gzip_compress = sub_row_21[-2]
    gzip_uncompress = sub_row_22[-2]

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

    my_worksheet.write('A5', my_file)
    my_worksheet.write('B5', 'zstd_uncompress_legacy')
    my_worksheet.write('C5', zstd_uncompress_legacy)

    my_worksheet.write('A6', my_file)
    my_worksheet.write('B6', 'random_read')
    my_worksheet.write('C6', random_read)

    my_worksheet.write('A7', my_file)
    my_worksheet.write('B7', 'black_scholes_serial')
    my_worksheet.write('C7', black_scholes_serial)

    my_worksheet.write('A8', my_file)
    my_worksheet.write('B8', 'string_search')
    my_worksheet.write('C8', string_search)

    my_worksheet.write('A9', my_file)
    my_worksheet.write('B9', 'random_write')
    my_worksheet.write('C9', random_write)

    my_worksheet.write('A10', my_file)
    my_worksheet.write('B10', 'object_detection')
    my_worksheet.write('C10', object_detection)

    my_worksheet.write('A11', my_file)
    my_worksheet.write('B11', 'ef_face_recognition')
    my_worksheet.write('C11', ef_face_recognition)

    my_worksheet.write('A12', my_file)
    my_worksheet.write('B12', 'zstd_compress_legacy')
    my_worksheet.write('C12', zstd_compress_legacy)

    my_worksheet.write('A13', my_file)
    my_worksheet.write('B13', 'zstd_uncompress_streaming')
    my_worksheet.write('C13', zstd_uncompress_streaming)

    my_worksheet.write('A14', my_file)
    my_worksheet.write('B14', 'fdt_by_medianflow_tracker')
    my_worksheet.write('C14', fdt_by_medianflow_tracker)

    my_worksheet.write('A15', my_file)
    my_worksheet.write('B15', 'black_scholes_parallel')
    my_worksheet.write('C15', black_scholes_parallel)

    my_worksheet.write('A16', my_file)
    my_worksheet.write('B16', 'create_sqlite_blob')
    my_worksheet.write('C16', create_sqlite_blob)

    my_worksheet.write('A17', my_file)
    my_worksheet.write('B17', 'video_colorization')
    my_worksheet.write('C17', video_colorization)

    my_worksheet.write('A18', my_file)
    my_worksheet.write('B18', 'external_sort')
    my_worksheet.write('C18', external_sort)

    my_worksheet.write('A19', my_file)
    my_worksheet.write('B19', 'chacha20_encrypt_openssl')
    my_worksheet.write('C19', chacha20_encrypt_openssl)

    my_worksheet.write('A20', my_file)
    my_worksheet.write('B20', 'colorization')
    my_worksheet.write('C20', colorization)

    my_worksheet.write('A21', my_file)
    my_worksheet.write('B21', 'hdr_stitch')
    my_worksheet.write('C21', hdr_stitch)

    my_worksheet.write('A22', my_file)
    my_worksheet.write('B22', 'aes_gcm_encrypt_mt')
    my_worksheet.write('C22', aes_gcm_encrypt_mt)

    my_worksheet.write('A23', my_file)
    my_worksheet.write('B23', 'memory_workload')
    my_worksheet.write('C23', memory_workload)

    my_worksheet.write('A24', my_file)
    my_worksheet.write('B24', 'chacha20_decrypt_openssl')
    my_worksheet.write('C24', chacha20_decrypt_openssl)

    my_worksheet.write('A25', my_file)
    my_worksheet.write('B25', 'gzip_compress')
    my_worksheet.write('C25', gzip_compress)

    my_worksheet.write('A26', my_file)
    my_worksheet.write('B26', 'gzip_uncompress')
    my_worksheet.write('C26', gzip_uncompress)

    
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
