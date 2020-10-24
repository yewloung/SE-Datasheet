import requests as req
from urllib.request import urlopen
from bs4 import BeautifulSoup
import csv
import re
import urllib3
import xlsxwriter
import pandas as pd

def tableDataText(table):
    """Parses a html segment started with tag <table> followed by multiple <tr> (table rows) and
    inner <td> (table data) tags. It returns a list of rows with inner columns.
    Accepts only one <th> (table header/data) in the first row.
    """
    def rowgetDataText(tr, coltag='td'): # td (data) or th (header)
        return [td.get_text(strip=True) for td in tr.find_all(coltag)]
    rows = []
    trs = table.find_all('tr')
    headerow = rowgetDataText(trs[0], 'th')

    if headerow: # if there is a header row include first
        rows.append(headerow)
        trs = trs[1:]
    for tr in trs: # for every table row
        rows.append(rowgetDataText(tr, 'td')) # data row
    return rows

reference_data = []
tab_data = []
section_data = []
parameter_data = []
value_data = []

headers = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/85.0.4183.121 '
                         'Safari/537.36 RuxitSynthetic/1.0 v6890906368 t38550 ath9b965f92 altpub cvcv=2'}

page = req.get('https://www.se.com/sg/en/product/RM17TE00/', headers=headers)

text_page = page.text  # convert page into text format
text_page = text_page.replace('</br>', ' | ')  # replace </br> text with (space)|(space)
#print(text_page)

soup = BeautifulSoup(text_page, 'html.parser')  # parse text according to html format

#head_tag = soup.head
#head_content = soup.contents
#head_content0 = soup.contents[0]
#print(len(list(head_tag)), len(list(head_content)), len(list(head_content0)))

#  Get product reference name
reference = soup.find('h1', {'data-autotests-id':'mobile-product-id'})
reference = reference.next
reference = reference.string.replace(" ", "")

htmltabs = soup.find_all('table', {'class':'pes-table'})

htmltables = soup.find_all('table', {'class':'pes-table'})
for htmltable in htmltables:
    #print(htmltable)  # all text content enclosed between <table class='pes-table'>...</table>

    tablenames = htmltable.find_all(['caption'])
    for tablename in tablenames:
        #print(tablename)  # all text content enclosed between <caption>...</caption>

        tablecontents = htmltable.find_all(['tbody'])
        for tablecontent in tablecontents:
            #print(tablecontent)  # all text content enclosed between <tbody>...</tbody>

            tablerows = htmltable.find_all(['tr'])
            for tablerow in tablerows:
                #print(tablerow)  # all text content enclosed between <tr>...</tr>

                reference_data.append(reference)
                section_data.append(' '.join(tablename.text.split()))

                tableheaders = tablerow.find_all(['th'])
                for tableheader in tableheaders:
                    #print(tableheader)  # all text content enclosed between <th>...</th>
                    parameter_data.append(' '.join(tableheader.text.split()))

                tabledatas = tablerow.find_all(['td'])
                for tabledata in tabledatas:
                    #print(tabledata)  # all text content enclosed between <td>...</td>
                    value_data.append(' '.join(tabledata.text.split()))

value_data[:] = [x for x in value_data if x]

print('reference:', len(reference_data), '; ',
      'Section:', len(section_data), '; ',
      'Parameters:', len(parameter_data), '; ',
      'Value:', len(value_data), ';',
      'tablerows:', len(tablerows), ';',
      'tableheaders:', len(tableheaders), ';',
      'tabledatas:', len(tabledatas)
      )

df = pd.DataFrame({'Reference': reference_data,
                   'Section': section_data,
                   'Parameters': parameter_data,
                   'Value': value_data})
print(df)
df.to_excel(r'SE_datasheet.xlsx', index=False, header=True)