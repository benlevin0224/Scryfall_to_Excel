import urllib.request, urllib.parse, urllib.error
import json
import ssl
from time import sleep
import xlsxwriter

# Ignore SSL certificate errors
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

#to create and write to CSV file
file_name = "test.xlsx"#input('enter your file name here. ')+".xlsx"
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

page_num = 1
serviceurl = "https://api.scryfall.com/cards/search?&page=%s\&q=" % page_num
card_type = "elf"#input("Enter card type here. All lower case. ")
cmdr = "G"#input("Enter commander's colors, WUBRG format. ")

url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
print("Retrieving", url)
uh = urllib.request.urlopen(url, context=ctx)
data = uh.read().decode()
js = json.loads(data)
has_more = js["has_more"]

headers = ["name","mana","oracle text"]
for col_num, data in enumerate(headers):
    worksheet.write(0, col_num, data)

column = 0
row = 1
while has_more is True:
    count = 0
    for cards in range(175):
        try:
            lst = [(js["data"][count]["name"]), (js["data"][count]["mana_cost"]), (js["data"][count]["oracle_text"])]
            for column, value in enumerate(lst):
                worksheet.write(row, column, value)
            row += 1
            count += 1
            sleep(0.1)
        except:
            column = 0
            worksheet.write(row,column,js["data"][count]["name"])
            row += 1
            card_side = 0
            for card_side in range(2):
                lst = [(js["data"][count]["card_faces"][card_side]["name"]),(js["data"][count]["card_faces"][card_side]["mana_cost"]),
                (js["data"][count]["card_faces"][card_side]["oracle_text"].encode("UTF-8").decode())]
                for column, value in enumerate(lst):
                    worksheet.write(row, column, value)
                    card_side += 1
                row += 1
            count += 1

    page_num += 1
    serviceurl = "https://api.scryfall.com/cards/search?&page=%s\&q=" % page_num
    url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
    uh = urllib.request.urlopen(url, context=ctx)
    data = uh.read().decode()
    js = json.loads(data)
    has_more = js["has_more"]
    print("Retrieving", url)

else:
    count = 0
    for cards in range((js["total_cards"]-(175*(page_num-1)))):
        try:
            lst = [(js["data"][count]["name"]), (js["data"][count]["mana_cost"]), (js["data"][count]["oracle_text"])]
            for column, value in enumerate(lst):
                worksheet.write(row, column, value)
            row += 1
            count += 1
            sleep(0.1)
        except:
            column = 0
            worksheet.write(row,column,js["data"][count]["name"])
            row += 1
            card_side = 0
            for card_side in range(2):
                lst = [(js["data"][count]["card_faces"][card_side]["name"]),(js["data"][count]["card_faces"][card_side]["mana_cost"]),
                (js["data"][count]["card_faces"][card_side]["oracle_text"].encode("UTF-8").decode())]
                for column, value in enumerate(lst):
                    worksheet.write(row, column, value)
                    card_side += 1
                row += 1
            count += 1

print('Finished exporting')
workbook.close()
