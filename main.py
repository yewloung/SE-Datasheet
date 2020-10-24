import requests as req
from bs4 import BeautifulSoup
import re
import xlsxwriter
import pandas as pd
import time
import os.path
import sys
import keyboard

# define list holders to store scraped data
reference_data = []
section_data = []
parameter_data = []
value_data = []
status = []
found_count = 0
notfound_count = 0
ave_execution_time = 0
ave_total_time = 0
ave_time_left = 0
input_file = r'ref.xlsx'
output_file = r'SE_web_datasheet.xlsx'

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
        text_page = text_page.replace('</br>', ' | ')  # replace </br> text with (space)|(space)

        # print(text_page)

        soup = BeautifulSoup(text_page, 'html.parser')  # parse text according to html format

        # head_tag = soup.head
        # head_content = soup.contents
        # head_content0 = soup.contents[0]
        # print(len(list(head_tag)), len(list(head_content)), len(list(head_content0)))

        #  Get product reference name
        reference = soup.find('h1', {'data-autotests-id': 'mobile-product-id'})
        reference = reference.next
        reference = reference.string.replace(" ", "")

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

    if os.path.isfile(input_file) == True:
        # read ref.xlsx, expect reference name must be put at col 1
        ref_df = pd.read_excel(input_file, header=None)

        if (ref_df.empty != True):
            ref_df[1] = 'https://www.se.com/sg/en/product/' + ref_df[0] + '/'  # create url for each references @ 2nd col
            urls = ref_df[1].tolist()  # put url into list
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

            # put the status of scraped reference in col 3
            ref_df[2] = status

            df.to_excel(output_file, index=False, header=True)  # store all scraped data into output_file
            ref_df.to_excel(input_file, index=False, header=None)  # update status of scraped status back to input_file
            #print(df)

            # get total program execution time
            print('\n--- Total program run time is %s seconds ---' % round(time.time() - start_time, 2))
            print('--- Total / found / not found :', len(urls), '/', found_count, '/', notfound_count, ' ---')

        else:
            print('\n')
            print('************************************************************************************')
            print('No data found in', input_file)
            print('Input your target commercial references in column A')
            print('Re-run the program once completed')
            print('************************************************************************************')
            exit()

    else:
        workbook = xlsxwriter.Workbook(input_file)
        worksheet = workbook.add_worksheet()
        workbook.close()
        print('\n')
        print('************************************************************************************')
        print(input_file, 'file not found.')
        print('Empty', input_file, 'is created. Input your target commercial references in column A')
        print('Re-run the program once completed')
        print('************************************************************************************')
        exit()

if __name__ == "__main__":
    # capture start of program execution time
    start_time = time.time()
    main()





