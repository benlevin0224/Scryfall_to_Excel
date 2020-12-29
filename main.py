from time import sleep
import xlsxwriter
import requests


def single_sided_cards(counts, rows):
    lst = [(js["data"][counts]["name"]),
           (js["data"][counts]["mana_cost"]),
           (js["data"][counts]["oracle_text"])]
    for column, value in enumerate(lst):
        worksheet.write(rows, column, value)
        sleep(.1)


def double_sided_cards(counts, rows, card_sides):
    lst = [
        (js["data"][counts]["card_faces"][card_sides]["name"]),
        (js["data"][counts]["card_faces"][card_sides]["mana_cost"]),
        (js["data"][counts]["card_faces"][card_sides]["oracle_text"])
    ]
    for column, value in enumerate(lst):
        worksheet.write(rows, column, value)
        sleep(.1)



file_name = input("enter your file name here. ") + ".xlsx"
card_type = input("Enter card type here. All lower case. ")
cmdr = input("Enter commander's colors, WUBRG format. ")

serviceurl = "https://api.scryfall.com/cards/search?&page=1&q="
url = f"{serviceurl}+type%3A{card_type}+id%3A{cmdr}"

# context manager to remove the close() statement
with xlsxwriter.Workbook(file_name) as workbook:
    worksheet = workbook.add_worksheet()
    headers = ["name", "mana", "oracle text"]
    for col_num, data in enumerate(headers):
        worksheet.write(0, col_num, data)

    row = 1
    while True:
        count = 0
        response = requests.get(url)
        js = response.json()
        print("Retrieving URL: " + url)
        for single_sided in range(len(js["data"])):
            try:
                single_sided_cards(count, row)
                row += 1
            except KeyError:
                card_side = 0
                worksheet.write(row, 0, js["data"][count]["name"])  # The 0 is column number
                row += 1
                for double_sided in range(2):
                    double_sided_cards(count, row, card_side)
                    row += 1
                    card_side += 1
            count += 1

        if "next_page" in js:
            url = js["next_page"]
        else:
            break
print("Finished exporting")