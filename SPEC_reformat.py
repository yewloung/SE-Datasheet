import re
import xlsxwriter
import pandas as pd
import time
import os.path
import sys
import keyboard

# define list holders to store scraped data
app_name = 'SPEC_reformat'
version = 'V1.0'
author = 'YL Liew'

reference_row = 3  # Row 1 = 0, Row 2 = 1, Row 3 = 2 ....
param_col = 5  # Col A = 0, Col B = 1, Col C = 2 ....
first_value_col = 9  # Col A = 0, Col B = 1, Col C = 2 ....
second_value_col = 3  # Col A = 0, Col B = 1, Col C = 2 ....
third_value_col = 2  # Col A = 0, Col B = 1, Col C = 2 ....
forth_value_col = 1  # Col A = 0, Col B = 1, Col C = 2 ....

param_id = []
value_data = []
status = []
valid_sheet = []

total_spec_count = 0
ave_execution_time = 0
ave_total_time = 0
ave_time_left = 0
input_file = r'product spec.xlsx'
output_file = r'SE_spec.xlsx'
sample_file = r'spec_format.xlsx'

#pd.set_option('display.max_rows', None)
#pd.set_option('display.max_columns', None)
#pd.set_option('display.width', None)
#pd.set_option('display.max_colwidth', -1)

def get_param_dict(sample_file):
    param_df = pd.read_excel(sample_file, sheet_name='param_lib', header=0)
    param_df.fillna(0, inplace=True)
    param_df.to_string()

    pivot = param_df.pivot_table(index=['ID', 'Param'],
                                 values=['Param'],
                                 aggfunc='count')

    pivot_REINDEX = pivot.reset_index()
    pivot_DICT = {k: v.tolist() for k, v in pivot_REINDEX.groupby('ID')['Param']}
    return pivot_DICT


def get_sample_spec():
    sample_spec_df = pd.read_excel(sample_file, header=None)
    print('---------------- Example of Original Product Specification Format You Should First Develop ----------------')
    print(sample_spec_df)
    print('\n')
    print('Note:')
    print('- Top horizontal number list = Column Index', '|', 'Left most vertical number list = Row Index')
    print('- e.g: row for the specification column title is at 3')
    print('- e.g: column for the specification parameter << param_col >> is at 5')
    print('- e.g: First specification value column << First_value_col >> is at 9')
    print('- e.g: Additional fix specification value column << second_value_col >> is at 3')
    print('- e.g: Additional fix specification value column << third_value_col >> is at 2')
    print('- e.g: Additional fix specification value column << forth_value_col >> is at 1')
    print('---------------------------------- End of Product Specification Format ------------------------------------')
    print('\n')

    print('###########################################################################################################')
    print(app_name, 'is used for parsing original product specification file format,', input_file)
    print('as per the example above into 2 worksheets with specific format recognized by')
    print('SE_Datasheet_Scrape program.')
    print('Worksheet <spec>: Col A = concatenate comm ref + parameters | Col B = spec value')
    print('Worksheet <spec2>: Col A (unique) = concatenate comm ref + parameters | Col B = spec value')
    print('Only worksheet <spec2> is used by SE_Datasheet_Scrape program to combined with the')
    print('scraped datasheet data from se.com into pivot format to faciliate comparison & checking.')
    print('\n')
    print('Note:')
    print('- continously hit <Enter> to let the program to run base on all default setting')
    print('- OR, you could input your choice of column and row of data of your original product specification to allow')
    print('the program to parse if your original product specification format is different from the example')
    print('\n')
    print('ver: ', version, ' ' * 72, 'developed by: ', author)
    print('###########################################################################################################')
    to_continue = input('Hit <Enter> to start the program, character < e > to exit ..... ')
    while True:
        if to_continue == '':
            print('\n')
            break
        elif to_continue == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()
        else:
            to_continue = input('Hit <Enter> to start the program, character < e > to exit ..... ')


def get_user_input():
    global reference_row
    global param_col
    global first_value_col
    global second_value_col
    global third_value_col
    global forth_value_col

    while True:
        num = input('specification column title row [hit <Enter> for default value 3, character < e > to exit]: ')
        if num.isdigit():
            reference_row = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()

    while True:
        num = input('specification parameter << param_col >> [hit <Enter> for default value 5, character < e > to exit]: ')
        if num.isdigit():
            param_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()

    while True:
        num = input('First specification value column << First_value_col >> [hit <Enter> for default value 9, character < e > to exit]: ')
        if num.isdigit():
            first_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()

    while True:
        num = input('Fix specification value column << second_value_col >> [hit <Enter> for default value 3, character < e > to exit]: ')
        if num.isdigit():
            second_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()

    while True:
        num = input('Fix specification value column << third_value_col >> [hit <Enter> for default value 2, character < e > to exit]: ')
        if num.isdigit():
            third_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()

    while True:
        num = input('Fix specification value column << forth_value_col >> [hit <Enter> for default value 1, character < e > to exit]: ')
        if num.isdigit():
            forth_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()


def main():
    global ave_execution_time
    global ave_total_time
    global ave_time_left
    global total_spec_count

    global reference_row
    global param_col
    global first_value_col
    global second_value_col
    global third_value_col
    global forth_value_col

    global param_id
    global value_data
    global status
    global valid_sheet

    if os.path.isfile(input_file) == True:
        df = pd.ExcelFile(input_file)  # read product spec.xlsx
        param_DICT = get_param_dict(sample_file)  # convert spec parameters to dictionary key

        # dictionary holder to hold data read from product spec.xlsx
        sheet_to_df_map = {}
        new_sheet_to_df_map = {}  # not used

        # dataframe holder to hold the processed Param_ID, Value from product spec.xlsx
        temp_df = pd.DataFrame(columns=['Param_ID', 'Value'])
        final_df = pd.DataFrame(columns=['Param_ID', 'Value'])

        # check if all the worksheets in product spec.xlsx are empty
        # if all are empty, record the count of filled worksheet into variable, count_filled_sheet = 0
        count_filled_sheet = 0
        for sheet_name in df.sheet_names:
            # parse every filled worksheets in product spec.xlsx into dictionary
            sheet_to_df_map[sheet_name] = df.parse(sheet_name, header=None)
            if sheet_to_df_map[sheet_name].empty != True:
                count_filled_sheet = count_filled_sheet + 1

        # if count_filled_sheet = 0, then go to print message and abort program without further processing
        if count_filled_sheet > 0:
            for sheet_name in df.sheet_names:
                # parse every filled worksheets in product spec.xlsx into dictionary
                sheet_to_df_map[sheet_name] = df.parse(sheet_name, header=None)
                #sheet_to_df_map[sheet_name] = sheet_to_df_map[sheet_name].applymap(str)  # convert whole dataframe to string

                # ensure 'reference row' will not exceed the no. of rows of data frame
                if (sheet_to_df_map[sheet_name].shape[0] - 1) < reference_row:
                    Temp_reference_row = sheet_to_df_map[sheet_name].shape[0] - 1
                else:
                    Temp_reference_row = reference_row

                # ensure 'first_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < first_value_col:
                    Temp_first_value_col = sheet_to_df_map[sheet_name].shape[1] - 1
                else:
                    Temp_first_value_col = first_value_col

                # get the new col header content base on selected row, reference_row
                temp_headers = list(sheet_to_df_map[sheet_name].iloc[Temp_reference_row])

                # fill the new empty col header content with col? format to ensure there are not empty string
                for index, str_text in enumerate(temp_headers):
                    if str_text != str_text:
                        temp_headers[index] = 'col' + str(index)

                sheet_to_df_map[sheet_name].columns = temp_headers  # set new col header

                # reset the dataframe ignoring row before the selected row of original data frame for its header content
                sheet_to_df_map[sheet_name] = sheet_to_df_map[sheet_name][Temp_reference_row + 1:]
                sheet_to_df_map[sheet_name].reset_index(drop=True, inplace=True)
                #sheet_to_df_map.update({sheet_name: sheet_to_df_map[sheet_name]})

                # when all the columns' name before the defined 1st column data are not contain "col" wording,
                # it is considered valid data frame
                partial_temp_header = list(sheet_to_df_map[sheet_name].columns[0:Temp_first_value_col])
                partial_temp_header = list(map(lambda x: str(x), partial_temp_header))
                if sum(1 for s in partial_temp_header if 'col' in s) == 0:
                    valid_sheet.append(sheet_name)

            print('\n-------------------------------------------------------------------------------------------------')
            print('Valid Column -->', valid_sheet)
            print('reference_row -->', reference_row)
            print('param_col -->', param_col)
            print('first_value_col -->', first_value_col)
            print('second_value_col -->', second_value_col)
            print('third_value_col -->', third_value_col)
            print('forth_value_col -->', forth_value_col)
            print('\n-------------------------------------------------------------------------------------------------')

            ''' ------------------------------  create specification look-up format   ------------------------------ '''
            final_count = 0
            for sheet_name in valid_sheet:
                col_count = sheet_to_df_map[sheet_name].shape[1] - first_value_col
                row_count = sheet_to_df_map[sheet_name].shape[0]
                total_count = col_count * row_count
                final_count = final_count + total_count
                total_spec_count = final_count

            for sheet_name in valid_sheet:
                # ensure 'reference row' will not exceed the no. of rows of data frame
                if (sheet_to_df_map[sheet_name].shape[0] - 1) < reference_row:
                    reference_row = sheet_to_df_map[sheet_name].shape[0] - 1

                # ensure 'first_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < first_value_col:
                    first_value_col = sheet_to_df_map[sheet_name].shape[1] - 1

                # ensure 'param_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < param_col:
                    param_col = sheet_to_df_map[sheet_name].shape[1] - 1

                # ensure 'second_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < second_value_col:
                    second_value_col = sheet_to_df_map[sheet_name].shape[1] - 1

                # ensure 'third_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < third_value_col:
                    third_value_col = sheet_to_df_map[sheet_name].shape[1] - 1

                # ensure 'forth_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < forth_value_col:
                    forth_value_col = sheet_to_df_map[sheet_name].shape[1] - 1

                sheet_to_df_map[sheet_name] = sheet_to_df_map[sheet_name].fillna('')
                for i in range(first_value_col, sheet_to_df_map[sheet_name].shape[1]):
                    value_data = sheet_to_df_map[sheet_name].astype(str).iloc[:, i] + ' | ' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, second_value_col] + ',' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, third_value_col] + ',' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, forth_value_col]
                    #sheet_to_df_map[sheet_name].iloc[:, i] = value_data

                    ref = str(sheet_to_df_map[sheet_name].columns[i])
                    for y in range(sheet_to_df_map[sheet_name].shape[0]):
                        param = sheet_to_df_map[sheet_name].astype(str).iloc[y, param_col]

                        # convert actual param string to param dictionary key
                        for key, value in param_DICT.items():
                            for j in range(len(value)):
                                if param == value[j]:
                                    param = '_' + str(key)

                        param_id.append(ref + param)
                        final_count = final_count - 1
                        print(final_count, ref + param)

                    temp_df['Param_ID'] = pd.Series(param_id).astype(str)
                    temp_df['Value'] = value_data.astype(str)
                    final_df = final_df.append(temp_df, ignore_index=True)  # Final sorted look up specification

                    param_id.clear()  # clear dataframe for looping reuse

            ''' ---------------------------------------  Clean Data frame   ---------------------------------------- '''
            final_df['Value'] = final_df['Value'].str.replace(r'(\s[|]\s,,)', '')  # replace ' | ,,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'(,,)', '')  # replace ',,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'^(\s[|]\s,)', '')  # replace ' | ,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'^(\s[|]\s)', '')  # replace ' | ' pattern with nothing
            final_df2 = final_df.groupby(['Param_ID'])['Value'].apply('\n'.join).reset_index()

            ''' -----------------------------  export data frame to excel operation   ------------------------------ '''
            writer = pd.ExcelWriter(output_file,
                                    engine='xlsxwriter',
                                    options={'strings_to_urls': False,
                                             'strings_to_formulas': False,
                                             'strings_to_numbers': False}
                                    )  # associated panda to xlsxwriter engine

            # create 'spec' worksheet to hold all spec data for all references facilitating look up
            final_df.to_excel(writer, index=False, header=True, sheet_name='spec')
            final_df2.to_excel(writer, index=False, header=True, sheet_name='spec2')

            # create respective 'spec' worksheet to hold all spec data for checking
            for sheet_name in valid_sheet:
                sheet_to_df_map[sheet_name].to_excel(writer, index=False, header=True, sheet_name=sheet_name)
                #print(sheet_name)

            # assign exported datasheet workbook variable name as "workbook"
            workbook = writer.book

            # setup format condition to be used
            text_align_format = workbook.add_format()  # Add text alignment format
            text_align_format.set_text_wrap(True)
            text_align_format.set_align('top')
            text_align_format.set_align('left')

            # assign worksheet "spec" variable name as "spec_worksheet"
            spec_worksheet = writer.sheets['spec']
            spec_worksheet.set_column('A:A', 50, text_align_format)
            spec_worksheet.set_column('B:B', 150, text_align_format)
            spec_worksheet.freeze_panes(1, 1)
            spec_worksheet.set_zoom(75)

            # assign worksheet "spec" variable name as "spec2_worksheet"
            spec2_worksheet = writer.sheets['spec2']
            spec2_worksheet.set_column('A:A', 50, text_align_format)
            spec2_worksheet.set_column('B:B', 150, text_align_format)
            spec2_worksheet.freeze_panes(1, 1)
            spec2_worksheet.set_zoom(75)

            writer.save()

            # get total program execution time
            print('\n--- Total program run time is %s seconds ---' % round(time.time() - start_time, 2))
            print('--- Total captured specification :', total_spec_count, ' ---')

        else:
            print('\n')
            print('***************************************************************************************************')
            print('')
            print('No data found in', input_file)
            print('Input respective reference specification into', input_file)
            print('Re-run the program once completed')
            print('')
            print('************************************ Windows Close in 5 Seconds ***********************************')
            time.sleep(5)
            exit()

    else:
        workbook = xlsxwriter.Workbook(input_file)
        worksheet = workbook.add_worksheet('specification')
        workbook.close()
        print('\n')
        print('***************************************************************************************************')
        print('')
        print(input_file, 'file not found.')
        print('Please put', input_file, 'into same folder location of this program.')
        print('Re-run the program once completed')
        print('')
        print('************************************ Windows Close in 5 Seconds ***********************************')
        time.sleep(5)
        exit()

if __name__ == "__main__":
    # capture start of program execution time
    start_time = time.time()

    get_sample_spec()
    get_user_input()
    main()