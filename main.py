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
file_name = input('enter your file name here. ')+".xlsx"
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

page_num = 1
serviceurl = "https://api.scryfall.com/cards/search?&page=%s\&q=" % page_num
card_type = input("Enter card type here. All lower case. ")
cmdr = input("Enter commander's colors, WUBRG format. ")

url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
print("Retrieving", url)
uh = urllib.request.urlopen(url, context=ctx)
data = uh.read().decode()
js = json.loads(data)
has_more = js["has_more"]

headers = ["name","mana","oracle text"]
for col_num, data in enumerate(headers):
    worksheet.write(0, col_num, data)

row = 1
while has_more is True:
    count = 0
    column = 0
    for cards in range(175):
        try:
            name = js["data"][count]["name"]
            mana = js["data"][count]["mana_cost"]
            oracle_text = js["data"][count]["oracle_text"]
            worksheet.write(row,column,name)
            column += 1
            worksheet.write(row,column,mana)
            column += 1
            worksheet.write(row,column,oracle_text)
            column =0
            row += 1
            count += 1
            sleep(0.1)
        except:
            worksheet.write(row,column,js["data"][count]["name"])
            row += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["name"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["mana_cost"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["oracle_text"].encode("UTF-8").decode())
            row += 1
            column = 0
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["name"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["mana_cost"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["oracle_text"].encode("UTF-8").decode())
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
        try: #Need to fix for more than 1 page
            name = js["data"][count]["name"]
            mana = js["data"][count]["mana_cost"]
            oracle_text = js["data"][count]["oracle_text"]
            worksheet.write(row,column,name)
            column += 1
            worksheet.write(row,column,mana)
            column += 1
            worksheet.write(row,column,oracle_text)
            column =0
            row += 1
            count += 1
            sleep(0.1)
        except:
            worksheet.write(row,column,js["data"][count]["name"])
            row += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["name"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["mana_cost"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][0]["oracle_text"].encode("UTF-8").decode())
            row += 1
            column = 0
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["name"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["mana_cost"])
            column += 1
            worksheet.write(row,column,js["data"][count]["card_faces"][1]["oracle_text"].encode("UTF-8").decode())
            count += 1
workbook.close()
