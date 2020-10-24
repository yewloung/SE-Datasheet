from bs4 import BeautifulSoup

htmlText = """<tr>
  <td>
    This
    <a class="tip info" href="blablablablabla">is a first</a>
    sentence
    <br>
    This
    <a class="tip info" href="blablablablabla">is a second</a>
    sentence
    <br>This
    <a class="tip info" href="blablablablabla">is a third</a>
    sentence
    <br>
  </td>
</tr>"""

# these two steps are to put everything into one line. may not be necessary for you
htmlText = htmlText.replace("\n", " ")
while "  " in htmlText:
    htmlText = htmlText.replace("  ", " ")

# import into bs4
soup = BeautifulSoup(htmlText, "lxml")

# using https://stackoverflow.com/a/34640357/5702157
for br in soup.find_all("br"):
    br.replace_with("\n")

parsedText = soup.get_text()
while "\n " in parsedText:
    parsedText = parsedText.replace("\n ", "\n") # remove spaces at the start of new lines
print(parsedText.strip())

