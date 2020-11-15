import requests as req
from bs4 import BeautifulSoup
import validators
import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
import time
import os.path
import sys
import keyboard
import styleframe
from collections import Counter
#from NLP_Modules import get_text_similiarity

# define list holders to store scraped data
app_name = 'SE_Datasheet_Scrape'
version = 'V2.0'
author = 'YL Liew'
default_url_format = 'https://www.se.com/ww/en/product/<ref>/'
additional_url_format = ['https://www.se.com/us/en/product/<ref>/',
                         'https://www.se.com/in/en/product/<ref>/',
                         'https://www.se.com/fr/fr/product/<ref>/',
                         'https://www.schneider-electric.cn/zh/product/<ref>/']

other_urls = []
reference_data = []
range_data = []
section_data = []
parameter_data = []
value_data = []
param_id = []
run_ref = []
run_url = []
status = []
found_count = 0
notfound_count = 0
ave_execution_time = 0
ave_total_time = 0
ave_time_left = 0
input_file = r'ref.xlsx'
output_file = r'SE_web_datasheet.xlsx'
spec_file = r'SE_spec.xlsx'
spec_worksheet = 'spec2'
sample_file = r'spec_format.xlsx'
similiarity_TH = 20

# function to deal with excel formating
def autosize_excel_columns_df(worksheet, df, offset=0):
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 1
        worksheet.set_column(idx + offset, idx + offset, max_len)

def autosize_excel_columns(worksheet, df):
    autosize_excel_columns_df(worksheet, df.index.to_frame())
    autosize_excel_columns_df(worksheet, df, offset=df.index.nlevels)

# Program to find most frequent element in a list
def most_frequent(List):
    occurence_count = Counter(List)
    return occurence_count.most_common(1)[0][0]

# function to get unique values
def unique(listArray):
    # intilize a null list
    unique_list = []

    # traverse for all elements
    for x in listArray:
        # check if exists in unique_list or not
        if x not in unique_list:
            unique_list.append(x)
            # print list
    return unique_list

def get_startup_message():
    if (os.path.isfile(spec_file) == True):
        # read SE_spec.xlsx
        #spec_df = pd.read_excel(spec_file, sheet_name=spec_worksheet)
        print('###################################################################################################')
        print(app_name, 'is used for scraping datasheet information from the se.com')
        print('using default https://www.se.com/ww/en/product/<ref> url format & specific countries SE site: US,')
        print('India, France & China.')
        print('You could change to your other specific list of url format after hitting <Enter>,')
        print('allowing the program to apply your list in case datasheet is not found in the default url.')
        print('After performing web scraping operating, product specification will be pivoted into scraped ')
        print('datasheet to facilitate comparison & checking.')
        print('\n')
        print('ver: ', version, ' ' * 64, 'developed by: ', author)
        print('###################################################################################################')
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
    else:
        print('###################################################################################################')
        print(app_name, 'is used for scraping datasheet information from the se.com')
        print('using default https://www.se.com/ww/en/product/<ref> url format & specific countries SE site: US,')
        print('India, France & China.')
        print('You could change to your other specific list of url format after hitting <Enter>,')
        print('allowing the program to apply your list in case datasheet is not found in the default url.')
        print('Due to product specification,', spec_file, 'file is not found / created.')
        print('Hence, the program will only performing data scraping operation without pivoting')
        print('the product specification into scraped datasheet')
        print('\n')
        print('ver: ', version, ' ' * 64, 'developed by: ', author)
        print('###################################################################################################')
        to_continue = input('Hit <Enter> to start the program .....')
        while True:
            if to_continue == '':
                print('\n')
                break
            else:
                to_continue = input('Hit <Enter> to start the program .....')

def get_other_url(default_url_format):
    other_url_list = [default_url_format]
    other_url_list.extend(additional_url_format)

    while True:
        other_url = input('\nKey in other URL, hit <Enter> to end your input, character < e > to exit: ')
        #valid_url = validators.url(other_url)

        if other_url == '':
            return other_url_list
        elif '<ref>' not in other_url:
            print('No <ref> place folder for commercial reference found in your string,', other_url, 'is ignored.')
        elif other_url[0:8] != 'https://':
            other_url = 'https://' + other_url
            other_url_list.append(other_url)
        elif other_url == 'e':
            print('Program is aborting in 5 seconds.....')
            time.sleep(5)
            exit()
        else:
            other_url_list.append(other_url)

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

# function to scrape the web data base on given url
def get_web_datasheet(url):
    result = 'Found'
    headers = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                             'AppleWebKit/537.36 (KHTML, like Gecko) '
                             'Chrome/85.0.4183.121 '
                             'Safari/537.36 RuxitSynthetic/1.0 v6890906368 t38550 ath9b965f92 altpub cvcv=2'}

    try:
        page = req.get(url, headers=headers, timeout=5)

        text_page = page.text  # convert page into text format
        text_page = text_page.replace('</br>', '|')  # replace </br> text with |

        # print(text_page)

        soup = BeautifulSoup(text_page, 'html.parser')  # parse text according to html format

        # head_tag = soup.head
        # head_content = soup.contents
        # head_content0 = soup.contents[0]
        # print(len(list(head_tag)), len(list(head_content)), len(list(head_content0)))

        #  Get product reference name
        reference = soup.find('div', {'data-autotests-id': 'product-id'})
        reference = reference.next
        reference = reference.string.replace(" ", "")
        reference = reference.strip()

        htmltables = soup.find_all('table', {'class': 'pes-table'})
        for htmltable in htmltables:
            # print(htmltable)  # all text content enclosed between <table class='pes-table'>...</table>

            tablenames = htmltable.find_all(['caption'])
            for tablename in tablenames:
                # print(tablename)  # all text content enclosed between <caption>...</caption>

                tablecontents = htmltable.find_all(['tbody'])
                for tablecontent in tablecontents:
                    # print(tablecontent)  # all text content enclosed between <tbody>...</tbody>

                    tablerows = htmltable.find_all(['tr'])
                    for tablerow in tablerows:
                        # print(tablerow)  # all text content enclosed between <tr>...</tr>

                        reference_data.append(reference)
                        section_data.append(' '.join(tablename.text.split()))

                        tableheaders = tablerow.find_all(['th'])
                        for tableheader in tableheaders:
                            # print(tableheader)  # all text content enclosed between <th>...</th>
                            parameter_data.append(' '.join(tableheader.text.split()))

                        tabledatas = tablerow.find_all(['td'])
                        for tabledata in tabledatas:
                            # print(tabledata)  # all text content enclosed between <td>...</td>
                            value_data.append(' '.join(tabledata.text.split()))

        value_data[:] = [x for x in value_data if x]
    except:
        result = 'Not Found'
    return result

def export_result_to_excel():
    pass

def main():
    global found_count
    global notfound_count
    global ave_execution_time
    global ave_total_time
    global ave_time_left
    global input_file
    global output_file
    global spec_file
    global spec_worksheet
    global version
    global similiarity_TH

    if (os.path.isfile(input_file) == True):
        # read ref.xlsx, expect reference name must be put at col 1
        ref_df = pd.read_excel(input_file, header=None)
        ref_status_df = pd.DataFrame()
        param_DICT = get_param_dict(sample_file)  # convert spec parameters to dictionary key

        count_scraped_url = 0
        if (ref_df.empty != True):
            refs = ref_df[0].astype(str)  # put all the commercial reference into list
            for i, ref in enumerate(refs):
                for j, other_url in enumerate(other_urls):
                    count_scraped_url = count_scraped_url + 1
                    url = other_url.replace('<ref>', ref)
                    result = get_web_datasheet(url)
                    run_ref.append(ref)
                    run_url.append(url)
                    status.append(result)

                    ave_execution_time = round(time.time() - start_time)
                    ave_total_time = (ave_execution_time / count_scraped_url) * ((len(refs) - 1 - i) + count_scraped_url)
                    ave_time_left = round(ave_total_time - ave_execution_time, 0)
                    print('[', ave_time_left, 'sec left ] ', i + 1, '/', len(refs), ': ', url, ' - ', result)

                    if result == 'Found':
                        found_count = found_count + 1
                        break
                    else:
                        notfound_count = notfound_count + 1

            ref_status_df[0] = run_ref
            ref_status_df[1] = run_url
            ref_status_df[2] = status

            # put all scraped data into data frame
            df = pd.DataFrame({'Reference': reference_data,
                               'Section': section_data,
                               'Parameters': parameter_data,
                               'Value': value_data})

            # remove newline character in front of the text
            df['Reference'] = df['Reference'].replace('\n', ' ', regex=True)

            # split string into multiple lines when detecting "|" character
            df['Value'] = df['Value'].str.replace('|', '\n')

            # add param id information
            df['Temp_Parameters'] = df['Parameters']
            for i in range(len(df)):
                # convert actual param string to param dictionary key
                for key, value in param_DICT.items():
                    for j in range(len(value)):
                        if df.loc[i, 'Temp_Parameters'] == value[j]:
                            df.loc[i, 'Temp_Parameters'] = '_' + str(key)

            df['Param_ID'] = df['Reference'] + df['Temp_Parameters']
            # df['Param_ID'] = df['Reference'] + df['Parameters']

            if (os.path.isfile(spec_file) == True):
                spec_df = pd.read_excel(spec_file, sheet_name=spec_worksheet)
                # map spec data
                # df['Spec_ID'] = df['Param_ID'].map(spec_df.set_index('Param_ID')['Param_ID'])
                df['Spec_Data'] = df['Param_ID'].map(spec_df.set_index('Param_ID')['Value']).astype(str)

                # check data sheet value vs spec value similiarity
                print('\n--- Compare similarity of tabulated datasheet value vs specification value ---')
                for index, row in df.iterrows():
                    #percent_similiarity = get_text_similiarity(row['Spec_Data'], row['Value'], 'xxxxx')
                    #df.loc[index, 'Similiarity'] = percent_similiarity[0]
                    df.loc[index, 'Similiarity'] = 0
                    #print(index, ':', row['Spec_Data'], ',', row['Value'], '-->', percent_similiarity[1])

                df['X_Correction'] = ''

            ''' --------------------------------------------------------------------------------------------------- '''
            # generate pivot table, put 'Value' for each 'Reference' arranged into column
            if (os.path.isfile(spec_file) == True):
                pivot = df.pivot_table(index=['Section', 'Parameters'],
                                       columns=['Reference'],
                                       values=['Value', 'Spec_Data', 'Similiarity', 'X_Correction'],
                                       aggfunc={'Value': lambda x: ' '.join(x),
                                                'Spec_Data': lambda x: ' '.join(x),
                                                'Similiarity': lambda x: x,
                                                'X_Correction': lambda x: x})
            else:
                pivot = df.pivot_table(index=['Section', 'Parameters'],
                                       columns=['Reference'],
                                       values=['Value'],
                                       aggfunc={'Value': lambda x: ' '.join(x)})

            pivot = pivot.swaplevel(0, 1, axis=1).sort_index(axis=1)  # swap column levels
            pivot.insert(loc=0, column='All_Cnt', value=10)
            pivot.insert(loc=1, column='Uniq_Cnt', value=100)
            pivot.insert(loc=2, column='Most_Cnt', value=200)
            pivot.insert(loc=3, column='Most_Freq_Val', value=pivot.index.get_level_values(1))

            # slide the pivot dataframe base on specified column
            #idx = pd.IndexSlice
            #all_values = pivot.loc[:, idx[:, 'Value']]
            #all_specs = pivot.loc[:, idx[:, 'Spec_Data']]

            ''' --------------------------------------------------------------------------------------------------- '''
            count_no_item_in_list = lambda u: len(u)
            count_no_of_unique_item_in_list = lambda v: len(v.unique())
            count_no_of_most_freq_item_in_list = lambda w: w.tolist().count(most_frequent(w))
            #count_no_of_most_freq_item_in_list = lambda w: w.count(most_frequent(w))
            #count_no_of_most_freq_item_in_list = lambda w: dict((i, w.count(i)) for i in set(w))
            join_all_item_in_list_byNewLine = lambda x: '\n'.join(x)
            join_all_unique_item_in_list_byNewLine = lambda y: '|\n'.join(map(str, unique(y)))
            most_freq_item_in_list = lambda z: most_frequent(z)

            count_no_item_in_list.__name__ = 'all_item_count'
            count_no_of_unique_item_in_list.__name__ = 'unique_item_count'
            count_no_of_most_freq_item_in_list.__name__ = 'most_freq_item_count'
            join_all_item_in_list_byNewLine.__name__ = 'all_item_list'
            join_all_unique_item_in_list_byNewLine.__name__ = 'all_unique_item_list'
            most_freq_item_in_list.__name__ = 'most_freq_item'

            list_of_funcs = [count_no_item_in_list,
                             count_no_of_unique_item_in_list,
                             count_no_of_most_freq_item_in_list,
                             join_all_item_in_list_byNewLine,
                             join_all_unique_item_in_list_byNewLine,
                             most_freq_item_in_list]

            pivot1 = df.pivot_table(index=['Section', 'Parameters'],
                                    values=['Value'],
                                    aggfunc=list_of_funcs
                                    )
            pivot1['tempAll_Cnt'] = pivot1['all_item_count']
            pivot1['tempUniq_Cnt'] = pivot1['unique_item_count']
            pivot1['tempMost_Cnt'] = pivot1['most_freq_item_count']
            pivot1['tempMost_Freq_Val'] = pivot1['most_freq_item']

            #pivot1 = df.pivot_table(index=['Section', 'Parameters'],
            #                        values=['Value'],
            #                        aggfunc={lambda v: most_frequent(v),
            #                                 lambda w: '\n'.join(w),
            #                                 lambda x: len(x.unique()),
            #                                 lambda y: len(y),
            #                                 lambda z: '|\n'.join(map(str, unique(z)))}
            #                        )

            ''' -------------------------  map pivot1 data to pivot   --------------------------- '''
            pivot['All_Cnt'] = pivot1['tempAll_Cnt']
            pivot['Uniq_Cnt'] = pivot1['tempUniq_Cnt']
            pivot['Most_Cnt'] = pivot1['tempMost_Cnt']
            pivot['Most_Freq_Val'] = pivot1['tempMost_Freq_Val']

            pivot1.drop(('tempAll_Cnt', ''), axis=1, inplace=True)  # drop multilevel indexed column
            pivot1.drop(('tempUniq_Cnt', ''), axis=1, inplace=True)  # drop multilevel indexed column
            pivot1.drop(('tempMost_Cnt', ''), axis=1, inplace=True)  # drop multilevel indexed column
            pivot1.drop(('tempMost_Freq_Val', ''), axis=1, inplace=True)  # drop multilevel indexed column

            ''' --------------------  export data frame to excel operation   -------------------- '''
            writer = pd.ExcelWriter(output_file,
                                    engine='xlsxwriter',
                                    options={'strings_to_urls': False,
                                             'strings_to_formulas': False,
                                             'strings_to_numbers': False}
                                    )  # associated panda to xlsxwriter engine

            # update status of web scrape to "status" worksheet
            ref_status_df.to_excel(writer, index=True, header=False, sheet_name='status')

            # export df data frame to "raw_datasheet" worksheet
            df.to_excel(writer, index=True, header=True, sheet_name='raw_datasheet')

            # export pivot data frame to "pivot_datasheet" worksheet
            pivot.to_excel(writer, index=True, header=True, sheet_name='pivot_datasheet')

            # export pivot1 data frame to "pivot1_datasheet" worksheet
            pivot1.to_excel(writer, index=True, header=True, sheet_name='pivot1_datasheet')

            if (os.path.isfile(spec_file) == True):
                # export spec_df data frame to "spec" worksheet
                spec_df.to_excel(writer, index=False, header=True, sheet_name='spec')

            # update status of scraped status back to input_file
            #ref_df.to_excel(input_file, index=False, header=None)

            # assign exported datasheet workbook variable name as "workbook"
            workbook = writer.book

            # setup format condition to be used
            text_align_format = workbook.add_format()  # Add text alignment format
            text_align_format.set_text_wrap(True)
            text_align_format.set_align('top')
            text_align_format.set_align('left')

            greenBG_whiteTEXT = workbook.add_format()  # Add color format green cell white text
            greenBG_whiteTEXT.set_font_color('white')
            greenBG_whiteTEXT.set_bg_color('green')

            yellowBG_blackTEXT = workbook.add_format()  # Add color format yellow cell black text
            yellowBG_blackTEXT.set_font_color('black')
            yellowBG_blackTEXT.set_bg_color('yellow')

            blueBG_whiteTEXT = workbook.add_format()  # Add color format blue cell white text
            blueBG_whiteTEXT.set_font_color('white')
            blueBG_whiteTEXT.set_bg_color('cyan')

            greyBG_whiteTEXT = workbook.add_format()  # Add color format blue cell white text
            greyBG_whiteTEXT.set_font_color('white')
            greyBG_whiteTEXT.set_bg_color('silver')

            redTEXT = workbook.add_format()  # Add color format red text
            redTEXT.set_font_color('red')

            addBORDER = workbook.add_format()  # Add border
            addBORDER.set_border(1)

            ''' --------------------------------------------------------------------------------------------------- '''
            # assign worksheet "status" variable name as "status_worksheet"
            status_worksheet = writer.sheets['status']
            autosize_excel_columns(status_worksheet, ref_df)
            status_worksheet.set_column(first_col=0, last_col=0, width=5)
            status_worksheet.set_column(first_col=1, last_col=1, width=20)
            status_worksheet.set_column(first_col=2, last_col=2, width=50)
            status_worksheet.set_column(first_col=3, last_col=3, width=15)
            status_worksheet.set_zoom(80)

            ''' --------------------------------------------------------------------------------------------------- '''
            # assign worksheet "raw_datasheet" variable name as "df_worksheet"
            df_worksheet = writer.sheets['raw_datasheet']
            df_worksheet.set_column('B:Z', 20, text_align_format)
            autosize_excel_columns(df_worksheet, df)
            df_worksheet.set_column(first_col=4, last_col=4, width=80, cell_format=text_align_format)
            df_worksheet.set_column(first_col=5, last_col=5, width=15, cell_format=text_align_format)
            df_worksheet.set_column(first_col=6, last_col=6, width=25, cell_format=text_align_format)
            df_worksheet.set_column(first_col=7, last_col=7, width=80, cell_format=text_align_format)
            df_worksheet.set_column(first_col=8, last_col=8, width=10, cell_format=text_align_format)
            # highlight cells color when similiarity exceed its defined threshold
            df_worksheet.conditional_format(1, 8, df.shape[0], 8,
                                            {'type': 'cell',
                                             'criteria': '>=',
                                             'value': similiarity_TH,
                                             'format': greenBG_whiteTEXT})
            df_worksheet.freeze_panes(1, 0)
            df_worksheet.set_zoom(80)
            df_worksheet.autofilter(0, 0, df.shape[0], df.shape[1])

            ''' --------------------------------------------------------------------------------------------------- '''
            # assign worksheet "pivot_datasheet" variable name as "pivot_worksheet"
            pivot_worksheet = writer.sheets['pivot_datasheet']

            # create list of columns with different width value to set columns width for multiple columns
            col_widths = {}
            col_widths[0] = 20
            col_widths[1] = 40
            col_widths[2] = 10
            col_widths[3] = 10
            col_widths[4] = 10
            col_widths[5] = 50

            for i in range(5, pivot.shape[1] + 2, 1):
                col_widths[i] = 50

            if (os.path.isfile(spec_file) == True):
                print('--- Compare similarity of pivoted datasheet value vs specification value ---')
                for i in range(6, pivot.shape[1] + 2, 4):  # specific set smaller col width for similarity col
                    col_widths[i] = 5
                    pivot_worksheet.conditional_format(2, i, df.shape[0], i,
                                                       {'type': 'cell',
                                                        'criteria': '>=',
                                                        'value': similiarity_TH,
                                                        'format': greenBG_whiteTEXT})

                print('--- Highlight cell color of datasheet value that appear the most ---')
                most_freq_list = pivot['Most_Freq_Val']
                #print(len(most_freq_list), pivot.shape[0], most_freq_list[0])

                # color non empty cells of all X_Correction columns as Red text
                for col in range(9, pivot.shape[1] + 2, 4):
                    pivot_worksheet.conditional_format(first_row=3, first_col=col,
                                                       last_row=pivot.shape[0] + 2, last_col=col,
                                                       options={'type': 'no_blanks', 'format': redTEXT})

                # color value of each ref which is same as most frequent value in yellow (when with SE_Spec)
                for col in range(8, pivot.shape[1] + 2, 4):  # start from col 8, step of 4 cols
                    for row in range(3, pivot.shape[0] + 3, 1):
                        ref_value = xl_rowcol_to_cell(row, 5, col_abs=True)  # convert cell coordinate to A1, A2...
                        comp_value = xl_rowcol_to_cell(row, col)  # convert cell coordinate to A1, A2...
                        criteria_syntax = '=(' + comp_value + '=' + ref_value + ')'
                        pivot_worksheet.conditional_format(first_row=row, first_col=col,
                                                           last_row=row, last_col=col,
                                                           options={'type': 'formula',
                                                                    'criteria': criteria_syntax,
                                                                    'format': yellowBG_blackTEXT})
            else:
                # color value of each ref which is same as most frequent value in yellow (when without SE_Spec)
                for col in range(6, pivot.shape[1] + 2, 1):  # start from col 6, step of 1 col
                    for row in range(3, pivot.shape[0] + 3, 1):
                        ref_value = xl_rowcol_to_cell(row, 5, col_abs=True)  # convert cell coordinate to A1, A2...
                        comp_value = xl_rowcol_to_cell(row, col)  # convert cell coordinate to A1, A2...
                        criteria_syntax = '=(' + comp_value + '=' + ref_value + ')'
                        pivot_worksheet.conditional_format(first_row=row, first_col=col,
                                                           last_row=row, last_col=col,
                                                           options={'type': 'formula',
                                                                    'criteria': criteria_syntax,
                                                                    'format': yellowBG_blackTEXT})

            # set all columns widths base on defined col_widths array defined
            for col_num, width in col_widths.items():
                pivot_worksheet.set_column(first_col=col_num, last_col=col_num, width=width,
                                           cell_format=text_align_format)

            # color empty cells of all columns to grey
            pivot_worksheet.conditional_format(first_row=3, first_col=6,
                                               last_row=pivot.shape[0] + 2, last_col=pivot.shape[1] + 1,
                                               options={'type': 'blanks', 'format': greyBG_whiteTEXT})

            # add border to the whole table
            pivot_worksheet.conditional_format(first_row=0, first_col=2,
                                               last_row=pivot.shape[0] + 2, last_col=pivot.shape[1] + 1,
                                               options={'type': 'no_errors', 'format': addBORDER})

            pivot_worksheet.freeze_panes(3, 6)
            pivot_worksheet.set_zoom(80)
            pivot_worksheet.autofilter(1, 0, pivot.shape[0], pivot.shape[1] + 1)
            pivot_worksheet.hide_gridlines(2)

            ''' --------------------------------------------------------------------------------------------------- '''
            # assign worksheet "pivot1_datasheet" variable name as "pivot1_worksheet"
            pivot1_worksheet = writer.sheets['pivot1_datasheet']
            pivot1_worksheet.set_column('A:A', 20, text_align_format)
            pivot1_worksheet.set_column('B:B', 40, text_align_format)
            pivot1_worksheet.set_column('C:C', 18, text_align_format)
            pivot1_worksheet.set_column('D:D', 18, text_align_format)
            pivot1_worksheet.set_column('E:E', 18, text_align_format)
            pivot1_worksheet.set_column(first_col=5, last_col=pivot1.shape[1] + 1, width=50, cell_format=text_align_format)
            pivot1_worksheet.freeze_panes(1, 2)
            pivot1_worksheet.set_zoom(80)
            pivot1_worksheet.autofilter(0, 0, pivot1.shape[0], pivot1.shape[1] + 1)

            ''' --------------------------------------------------------------------------------------------------- '''
            if (os.path.isfile(spec_file) == True):
                # assign worksheet "spec" variable name as "spec_worksheet"
                spec_worksheet = writer.sheets['spec']
                spec_worksheet.set_column('A:A', 50, text_align_format)
                spec_worksheet.set_column('B:B', 150, text_align_format)
                spec_worksheet.freeze_panes(1, 1)
                spec_worksheet.set_zoom(80)
                spec_worksheet.autofilter(0, 0, spec_df.shape[0], spec_df.shape[1])

            writer.save()

            ''' --------------------  read back worksheets created   -------------------- '''
            xls = pd.ExcelFile(output_file)

            # get total program execution time
            print('\n')
            print('---', xls.sheet_names, 'worksheets are created into', output_file, ' ---')
            print('--- Total program run time is %s seconds ---' % round(time.time() - start_time, 2))
            print('--- Total ref / found / not found / Total Trial:',
                  len(refs), '/',
                  found_count, '/',
                  len(refs) - found_count, '/',
                  found_count + notfound_count,
                  ' ---')

        else:
            print('\n')
            print('************************************************************************************')
            print('')
            print('No data found in', input_file)
            print('Input your target commercial references in column A of', input_file)
            print('Re-run the program once completed')
            print('')
            print('**************************** Windows Close in 5 Seconds ****************************')
            time.sleep(5)
            exit()

    else:
        workbook = xlsxwriter.Workbook(input_file)
        worksheet = workbook.add_worksheet()
        workbook.close()
        print('\n')
        print('************************************************************************************')
        print('')
        print(input_file, 'file not found.')
        print('Empty', input_file, 'is created into same folder location of this program.')
        print('Input your target commercial references in column A of', input_file)
        print('Re-run the program once completed')
        print('')
        print('**************************** Windows Close in 5 Seconds ****************************')
        time.sleep(5)
        exit()

if __name__ == "__main__":
    get_startup_message()
    other_urls = get_other_url(default_url_format)
    print('\n', 'List of trial URL: ', other_urls, '\n')

    # capture start of program execution time
    start_time = time.time()
    main()
