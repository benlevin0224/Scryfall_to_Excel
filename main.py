from time import sleep
import xlsxwriter
import requests

# to create and write to CSV file
file_name = input("enter your file name here. ") + ".xlsx"
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

serviceurl = "https://api.scryfall.com/cards/search?&page=1&q="
card_type = input("Enter card type here. All lower case. ")
cmdr = input("Enter commander's colors, WUBRG format. ")
url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
response = requests.get(url)
js = response.json()

headers = ["name", "mana", "oracle text"]
for col_num, data in enumerate(headers):
    worksheet.write(0, col_num, data)

column = 0
row = 1

while True:
    count = 0
    print("Retrieving URL: " + url)
    for cards in range(len(js["data"])):
        try:
            lst = [
                (js["data"][count]["name"]),
                (js["data"][count]["mana_cost"]),
                (js["data"][count]["oracle_text"])
            ]
            for column, value in enumerate(lst):
                worksheet.write(row, column, value)
            row += 1
            count += 1
            sleep(0.1)
        except KeyError:
            column = 0
            worksheet.write(row, column, js["data"][count]["name"])
            row += 1
            card_side = 0
            for card_side in range(2):
                lst = [
                    (js["data"][count]["card_faces"][card_side]["name"]),
                    (js["data"][count]["card_faces"][card_side]["mana_cost"]),
                    (js["data"][count]["card_faces"][card_side]["oracle_text"])
                ]
                for column, value in enumerate(lst):
                    worksheet.write(row, column, value)
                    card_side += 1
                row += 1
            count += 1
    try:
        url = js["next_page"]
        response = requests.get(url)
        js = response.json()
    except KeyError:
        break
print("Finished exporting")
workbook.close()
