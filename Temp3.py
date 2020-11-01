

#url_format = 'https://www.se.com/ww/en/product/<ref>'
url_format = 'https://www.se.com/ww/en/<ref>product/<ref>'

ref = 'RE17RAMU'

url = url_format.replace('<ref>', ref) + '/'
#url = url_format[0:8]

print(url)
print('cm' not in url)