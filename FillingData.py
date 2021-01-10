import openpyxl
import urllib.request
import json

# Enter both the file name and the sheet name
file_name = "Translated_Fiction.xlsx"
wb = openpyxl.load_workbook(file_name)
sheet = wb.get_sheet_by_name('Sheet1')

base_api_link = "https://www.googleapis.com/books/v1/volumes?q=isbn:"

# Enter starting cell number, usually 2
cellInt = 2
while True:
    isbnCell = 'C' + str(cellInt)
    titCell = 'D' + str(cellInt)
    autCell = 'E' + str(cellInt)
    pdCell = 'F' + str(cellInt)
    publisherCell = 'G' + str(cellInt)
    pubDateCell = 'H' + str(cellInt)


    isbnNumber = sheet[isbnCell].value

    if isbnNumber == None:
        break
    else:
        isbnNumber = str(isbnNumber)

    print()
    print(cellInt, end = ' ')

    with urllib.request.urlopen(base_api_link + isbnNumber) as f:
        text = f.read()

    decoded_text = text.decode("utf-8")
    obj = json.loads(decoded_text)

    # Code was put into a try and except block, as data was not usually found, returning a Key Error which would break the running code
    try:
        volume_info = obj["items"][0]

        authors = obj["items"][0]["volumeInfo"]["authors"]
        author = ",".join(authors)
        sheet[autCell] = author

        title = volume_info["volumeInfo"]["title"]
        sheet[titCell] = title

        description = volume_info["searchInfo"]["textSnippet"]
        description = str(description).replace("&#39;", "'").replace("&quot;--" or "&quot;", '"')
        sheet[pdCell] = description

        publisher = volume_info['volumeInfo']['publisher']
        sheet[publisherCell] = publisher

        pubDate = volume_info['volumeInfo']['publishedDate']
        sheet[pubDateCell] = pubDate

        print(author,title,description,publisher,pubDate,sep='\n')


    except KeyError:
        print('skipped',end = '')
        cellInt += 1
        continue

    cellInt += 1

wb.save(filename=file_name)

print()
print("Saved")
