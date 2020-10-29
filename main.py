import requests as req
from bs4 import BeautifulSoup
import re
import xlsxwriter
import pandas as pd
import time
import os.path
import sys
import keyboard
import styleframe

# define list holders to store scraped data
reference_data = []
range_data = []
section_data = []
parameter_data = []
value_data = []
param_id = []
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

# function to deal with excel formating
def autosize_excel_columns_df(worksheet, df, offset=0):
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 1
        worksheet.set_column(idx + offset, idx + offset, max_len)

def autosize_excel_columns(worksheet, df):
    autosize_excel_columns_df(worksheet, df.index.to_frame())
    autosize_excel_columns_df(worksheet, df, offset=df.index.nlevels)

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

    if (os.path.isfile(input_file) == True):
        # read ref.xlsx, expect reference name must be put at col 1
        ref_df = pd.read_excel(input_file, header=None)

        if (os.path.isfile(spec_file) == True):
            # read SE_spec.xlsx
            spec_df = pd.read_excel(spec_file, sheet_name=spec_worksheet)

        if (ref_df.empty != True):
            ref_df[1] = 'https://www.se.com/ww/en/product/' + ref_df[0] + '/'  # create url for each references @ 2nd col
            urls = ref_df[1].tolist()  # put url into list
            # https://www.se.com/ww/en/product/RM17TE00/
            # https://www.se.com/sg/en/product/RM17TE00/
            # print(urls)

            # loop through all urls to scrap all data into respective list holders
            for i, url in enumerate(urls):
                result = get_web_datasheet(url)
                status.append(result)

                if result == 'Found':
                    found_count = found_count + 1
                else:
                    notfound_count = notfound_count + 1

                ave_execution_time = round(time.time() - start_time)
                ave_total_time = (ave_execution_time / (i + 1)) * (len(urls))
                ave_time_left = round(ave_total_time - ave_execution_time, 0)
                print('[', ave_time_left, 'sec left ] ', i + 1, '/', len(urls), ': ', url, ' - ', result)

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
            df['Param_ID'] = df['Reference'] + df['Parameters']

            if (os.path.isfile(spec_file) == True):
                # map spec data
                # df['Spec_ID'] = df['Param_ID'].map(spec_df.set_index('Param_ID')['Param_ID'])
                df['Spec_Data'] = df['Param_ID'].map(spec_df.set_index('Param_ID')['Value']).astype(str)

            #df['Spec_ID'] = df['Param_ID']
            #df['Spec_Data'] = df['Value'] + 'XXXX'

            # put the status of scraped reference in col 3
            ref_df[2] = status

            # generate pivot table, put 'Value' for each 'Reference' arranged into column
            if (os.path.isfile(spec_file) == True):
                pivot = df.pivot_table(index=['Section', 'Parameters'],
                                       columns=['Reference'],
                                       values=['Value', 'Spec_Data'],
                                       aggfunc={'Value': lambda x: ' '.join(x),
                                                'Spec_Data': lambda x: ' '.join(x)})
            else:
                pivot = df.pivot_table(index=['Section', 'Parameters'],
                                       columns=['Reference'],
                                       values=['Value'],
                                       aggfunc={'Value': lambda x: ' '.join(x)})

            pivot = pivot.swaplevel(0, 1, axis=1).sort_index(axis=1)  # swap column levels

            pivot1 = df.pivot_table(index=['Section', 'Parameters'],
                                    values=['Value'],
                                    aggfunc=lambda x: '\n'.join(x))

            #pivot1 = df.pivot_table(index=['Section', 'Parameters', 'Value'],
            #                        values=['Value'],
            #                        aggfunc='count')

            #pivot1 = df.pivot_table(index=['Section', 'Parameters', 'Value'],
            #                        columns=['Reference'],
            #                        values=['Reference'],
            #                        aggfunc='count',
            #                        fill_value=0)

            #print(pivot1)

            ''' --------------------  export data frame to excel operation   -------------------- '''
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')  # associated panda to xlsxwriter engine

            # update status of web scrape to "status" worksheet
            ref_df.to_excel(writer, index=True, header=False, sheet_name='status')

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

            # assign worksheet "status" variable name as "status_worksheet"
            status_worksheet = writer.sheets['status']
            autosize_excel_columns(status_worksheet, ref_df)
            status_worksheet.set_zoom(80)

            # assign worksheet "raw_datasheet" variable name as "df_worksheet"
            df_worksheet = writer.sheets['raw_datasheet']
            df_worksheet.set_column('B:Z', 20, text_align_format)
            autosize_excel_columns(df_worksheet, df)
            #df_worksheet.set_column('E:E', 90, text_align_format)
            df_worksheet.set_column(first_col=4, last_col=4, width=90, cell_format=text_align_format)
            df_worksheet.set_column(first_col=5, last_col=5, width=40, cell_format=text_align_format)
            df_worksheet.set_column(first_col=6, last_col=6, width=40, cell_format=text_align_format)
            df_worksheet.set_column(first_col=7, last_col=7, width=95, cell_format=text_align_format)
            df_worksheet.freeze_panes(1, 0)
            df_worksheet.set_zoom(80)

            # assign worksheet "pivot_datasheet" variable name as "pivot_worksheet"
            pivot_worksheet = writer.sheets['pivot_datasheet']
            pivot_worksheet.set_column('A:A', 20, text_align_format)
            pivot_worksheet.set_column('B:B', 40, text_align_format)
            #autosize_excel_columns(pivot_worksheet, pivot)
            pivot_worksheet.set_column(first_col=2, last_col=(found_count + 1)*2, width=50, cell_format=text_align_format)
            pivot_worksheet.freeze_panes(3, 2)
            pivot_worksheet.set_zoom(80)

            # assign worksheet "pivot1_datasheet" variable name as "pivot1_worksheet"
            pivot1_worksheet = writer.sheets['pivot1_datasheet']
            pivot1_worksheet.set_column('A:A', 20, text_align_format)
            pivot1_worksheet.set_column('B:B', 40, text_align_format)
            #autosize_excel_columns(pivot1_worksheet, pivot1)
            pivot1_worksheet.set_column(first_col=2, last_col=found_count + 1, width=110, cell_format=text_align_format)
            pivot1_worksheet.freeze_panes(1, 2)
            pivot1_worksheet.set_zoom(80)

            if (os.path.isfile(spec_file) == True):
                # assign worksheet "spec" variable name as "spec_worksheet"
                spec_worksheet = writer.sheets['spec']
                spec_worksheet.set_column('A:A', 50, text_align_format)
                spec_worksheet.set_column('B:B', 150, text_align_format)
                # autosize_excel_columns(spec_worksheet, spec_df)
                spec_worksheet.freeze_panes(1, 1)
                spec_worksheet.set_zoom(80)

            writer.save()

            ''' --------------------  read back worksheets created   -------------------- '''
            xls = pd.ExcelFile(output_file)

            # get total program execution time
            print('\n')
            print('---', xls.sheet_names, 'worksheets are created into', output_file, ' ---')
            print('--- Total program run time is %s seconds ---' % round(time.time() - start_time, 2))
            print('--- Total / found / not found :', len(urls), '/', found_count, '/', notfound_count, ' ---')

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
    # capture start of program execution time
    start_time = time.time()
    main()





