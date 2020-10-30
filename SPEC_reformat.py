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

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)

def get_sample_spec():
    sample_spec_df = pd.read_excel(sample_file, header=None)
    print('------------------------ Example of Original Product Specification Format ---------------------------------')
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
    print('------------------------------ End of Product Specification Format ----------------------------------------')
    print('\n')

    print('#############################################################################################')
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
    print('ver: ', version, '                                                         ', 'developed by: ', author)
    print('#############################################################################################')
    to_continue = input('Hit <Enter> to start the program .....')
    while True:
        if to_continue == '':
            print('\n')
            break
        else:
            to_continue = input('Hit <Enter> to start the program .....')


def get_user_input():
    global reference_row
    global param_col
    global first_value_col
    global second_value_col
    global third_value_col
    global forth_value_col

    while True:
        num = input('specification column title row [hit <Enter> to accept default value 3]: ')
        if num.isdigit():
            reference_row = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            exit()

    while True:
        num = input('specification parameter << param_col >> [hit <Enter> to accept default value 5]: ')
        if num.isdigit():
            param_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            exit()

    while True:
        num = input('First specification value column << First_value_col >> [hit <Enter> to accept default value 9]: ')
        if num.isdigit():
            first_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            exit()

    while True:
        num = input('Fix specification value column << second_value_col >> [hit <Enter> to accept default value 3]: ')
        if num.isdigit():
            second_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            exit()

    while True:
        num = input('Fix specification value column << third_value_col >> [hit <Enter> to accept default value 2]: ')
        if num.isdigit():
            third_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
            exit()

    while True:
        num = input('Fix specification value column << forth_value_col >> [hit <Enter> to accept default value 1]: ')
        if num.isdigit():
            forth_value_col = int(num)
            break
        elif num == '':
            break
        elif num == 'e':
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

    if os.path.isfile(input_file) == True:
        df = pd.ExcelFile(input_file)  # read product spec.xlsx

        sheet_to_df_map = {}  # dictionary holder to hold data frames of each worksheet
        temp_df = pd.DataFrame(columns=['Param_ID', 'Value'])
        final_df = pd.DataFrame(columns=['Param_ID', 'Value'])

        count_filled_sheet = 0
        for sheet_name in df.sheet_names:
            # check every worksheet data format is a valid specification format
            # criteria: every columns headers before the 1st references specification value column MUST NOT empty
            # rows = df.shape[0] | cols = df.shape[1]
            sheet_to_df_map[sheet_name] = df.parse(sheet_name)  # convert every worksheet content into dictionary
            if sheet_to_df_map[sheet_name].empty != True:
                count_filled_sheet = count_filled_sheet + 1

        if count_filled_sheet > 0:
            for sheet_name in df.sheet_names:
                sheet_to_df_map[sheet_name] = df.parse(sheet_name)  # convert every worksheet content into dictionary

                # ensure defined 'reference row' will not exceed the no. of rows of data frame
                if (sheet_to_df_map[sheet_name].shape[0] - 1) < reference_row:
                    row_length = sheet_to_df_map[sheet_name].shape[0] - 1
                else:
                    row_length = reference_row - 1

                # ensure defined 'first_value_col' will not exceed the no. of cols of data frame
                if (sheet_to_df_map[sheet_name].shape[1] - 1) < first_value_col:
                    col_length = sheet_to_df_map[sheet_name].shape[1] - 1
                else:
                    col_length = first_value_col - 1

                # put row 3 as column name and remove row 1 & 2
                temp_header = sheet_to_df_map[sheet_name].iloc[row_length]
                for i in range(len(temp_header)):
                    if temp_header[i] != temp_header[i]:
                        temp_header[i] = 'col' + str(i)
                sheet_to_df_map[sheet_name].iloc[row_length] = temp_header
                sheet_to_df_map[sheet_name].columns = list(temp_header)
                sheet_to_df_map[sheet_name] = sheet_to_df_map[sheet_name][row_length + 1:]
                sheet_to_df_map[sheet_name].reset_index(drop=True, inplace=True)

                # when all the columns' name before the defined 1st column data are not contain "col" wording,
                # it is considered valid data frame
                partial_temp_header = sheet_to_df_map[sheet_name].columns[0:col_length]
                if sum(1 for s in partial_temp_header if 'col' in s) == 0:
                    valid_sheet.append(sheet_name)

                # print('Worksheet -->', sheet_name, '|',
                #      'Type -->', type(sheet_to_df_map[sheet_name]), '|',
                #      'Col / Column Qty -->', col_length, '/', sheet_to_df_map[sheet_name].shape[1],
                #      'Row / Row Qty -->', row_length, '/', sheet_to_df_map[sheet_name].shape[0],
                #      'Empty? -->', sheet_to_df_map[sheet_name].empty)
                # print(partial_temp_header)

            print('Valid Column -->', valid_sheet)

            ''' --------------------  create specification look-up format   -------------------- '''
            final_count = 0
            for sheet_name in valid_sheet:
                col_count = sheet_to_df_map[sheet_name].shape[1] - first_value_col
                row_count = sheet_to_df_map[sheet_name].shape[0] - 0
                total_count = col_count * row_count
                final_count = final_count + total_count
                total_spec_count = final_count

            for sheet_name in valid_sheet:
                sheet_to_df_map[sheet_name] = sheet_to_df_map[sheet_name].fillna('')
                for i in range(first_value_col, sheet_to_df_map[sheet_name].shape[1]):
                    value_data = sheet_to_df_map[sheet_name].astype(str).iloc[:, i] + ' | ' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, second_value_col] + ',' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, third_value_col] + ',' \
                                 + sheet_to_df_map[sheet_name].astype(str).iloc[:, forth_value_col]
                    #sheet_to_df_map[sheet_name].iloc[:, i] = value_data

                    ref = str(sheet_to_df_map[sheet_name].columns[i])
                    final_count = final_count - 1

                    for y in range(0, sheet_to_df_map[sheet_name].shape[0] - 1):
                        param = sheet_to_df_map[sheet_name].astype(str).iloc[y, param_col]
                        param_id.append(ref + param)
                        final_count = final_count - 1
                        print(final_count, ref + param)

                    temp_df['Param_ID'] = pd.Series(param_id)
                    temp_df['Value'] = value_data
                    final_df = final_df.append(temp_df, ignore_index=True)  # Final sorted look up specification
                    param_id.clear()

            ''' --------------------  Clean Data frame   -------------------- '''
            final_df['Value'] = final_df['Value'].str.replace(r'(\s[|]\s,,)', '')  # replace ' | ,,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'(,,)', '')  # replace ',,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'^(\s[|]\s,)', '')  # replace ' | ,' pattern with nothing
            final_df['Value'] = final_df['Value'].str.replace(r'^(\s[|]\s)', '')  # replace ' | ' pattern with nothing
            final_df2 = final_df.groupby(['Param_ID'])['Value'].apply('\n'.join).reset_index()

            ''' --------------------  export data frame to excel operation   -------------------- '''
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')  # associated panda to xlsxwriter engine

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
            print('************************************************************************************')
            print('')
            print('No data found in', input_file)
            print('Input respective reference specification into', input_file)
            print('Re-run the program once completed')
            print('')
            print('**************************** Windows Close in 5 Seconds ****************************')
            time.sleep(5)
            exit()

    else:
        workbook = xlsxwriter.Workbook(input_file)
        worksheet = workbook.add_worksheet('specification')
        workbook.close()
        print('\n')
        print('************************************************************************************')
        print('')
        print(input_file, 'file not found.')
        print('Please put', input_file, 'into same folder location of this program.')
        print('Re-run the program once completed')
        print('')
        print('**************************** Windows Close in 5 Seconds ****************************')
        time.sleep(5)
        exit()

if __name__ == "__main__":
    # capture start of program execution time
    start_time = time.time()

    get_sample_spec()
    get_user_input()
    main()