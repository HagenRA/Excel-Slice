# Aim: To read financial information from a source file (input) and write to a generic output file that allows for a
#      ppt file to have a static source to read for chart information. (This helps if excel chart uses =INDIRECT so that
#      you can enter the newly imported sheet name so that .

# Importing relevant modules that will be used in this program
import os
import pandas as pd

print('As a disclaimer, this excel reader-updater has the following restrictions and assumptions.')
print('\n1 - Source files and output files to are within the same directory as this script.')
print('2 - You have an output file which is NOT open during this process.')

print('\nThese are the excel files we have in this directory')
# Shows only .xlsx files
files = [f for f in os.listdir(os.curdir) if os.path.isfile(f) and f.endswith('.xlsx')]
files.sort()
index_files = {i: files[i] for i in range(0, len(files))}

check_open = ''
excel_select = ''
output_ID = ''
target_ID = ''

while excel_select.capitalize() != 'Y':
    for i in index_files:
        print(i, index_files[i])
    while output_ID == '':
        output_ID = input('Using the index number, which excel will we export to?: ')
        if output_ID.isalpha() or int(output_ID) > len(files):
            print('Please enter valid ID for export target.')
            output_ID = ''
    output = index_files[int(output_ID)]
    while target_ID == '':
        target_ID = input('Using the index number, which excel will we import from?: ')
        if target_ID.isalpha() or int(target_ID) > len(files):
            print('Please enter valid ID for import target.')
            target_ID = ''
    target = index_files[int(target_ID)]
    xls_file = pd.ExcelFile(target)
    export_file = pd.ExcelFile(output)
    print(f'\nThese are the sheets we can choose from the import target:')
    sheets = {i: xls_file.sheet_names[i] for i in range(0, len(xls_file.sheet_names))}
    for i in sheets:
        print(i, sheets[i])
    sheet_num = input('\nUsing the index number, which sheet are we importing?: ')
    active = sheets[int(sheet_num)]
    col_select = ''
    while col_select != 'Y':
        key_col = ''
        curr_col = ''
        # Ensures that that the input is valid.
        while True:
            key_col = input('\nWhat is the column used as the key?: ')
            if len(key_col) >= 1 and key_col.isalpha():
                break
            else:
                print('Please enter valid ID for key column.')
        while True:
            curr_col = input('What is the column of the current period?: ')
            if len(curr_col) >= 1 and curr_col.isalpha():
                break
            else:
                print('Please enter valid ID for current column.')
        # Defining number of columns to import
        col_num = 0
        num_test = False
        while num_test is False:
            col_num = input('Not including the key and current column, how many columns will we be importing today? ')
            try:
                col_num = int(col_num)
                num_test = True
            except:
                print('Insert a valid integer for number of columns to import.')
        print(f'We shall be importing {col_num} columns today.')
        # Running loops to add the target columns into the range between the key and current quarter
        inside = []
        # Validity test to ensure all column IDs are valid.
        for i in range(col_num):
            while True:
                x = input(f'The ID of column {i + 1} for the comparison is: ')
                if len(x) >= 1 and x.isalpha():
                    inside.append(x)
                    break
                else:
                    print('Please input a valid column ID.')
        # Now combining all the column IDs together
        comps = ''
        for i in inside:
            comps = comps + ', ' + i
        rng = str(f'{key_col.upper()}{comps.upper()}, {curr_col.upper()}')
        print(f'As such, we shall be importing these columns: Key=> {rng} <= Current Q.')
        col_select = input('Press [Y] to confirm these columns are correct. ').upper()
    print('\nPlease bear with the code as it is exporting it into a data frame so that it can be imported into excel.')
    # Writing to the data frame
    df = pd.read_excel(target, sheet_name=active, header=0, usecols=rng)
    # Sanity check that output file is closed before proceeding to avoid errors
    open_test = True
    while open_test is True:
        try:
            os.rename(output, 'tempfile.xls')
            os.rename('tempfile.xls', output)
            open_test = False
        except OSError:
            print(f'File is still open. Please close {output}')
            input(f'Press enter to confirm you have closed {output}')
    name_test = False
    while name_test is False:
        # To check if there will be duplicate sheets so that it won't have an export error
        sheet_ID = input(f'\nFor naming purposes, we shall be importing the {active} for what time period? ')
        sheet_ID = sheet_ID.upper()
        if len(active) >= 7:
            # Slicing the name down so that it doesn't become too long in the excel sheets
            active = active[:7]
            print(f'Name of original source sheet too long, renamed to {active}')
        sheet_name = sheet_ID + '-' + active
        name_test = True
        with pd.ExcelWriter(output, mode='a') as writer:
            df.to_excel(writer, sheet_name=sheet_ID + '-' + active)
        print(f'\nAdded {sheet_ID}-{active} to {output}.')

    excel_select = input('\nWill this be all the sheets you are importing today? [Y/N] ')

print(f'Export finished, now opening {output}')
os.system(f'start "EXCEL.EXE" "{output}" ')
