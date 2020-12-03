import urllib.request, urllib.parse, urllib.error
import json
import ssl
from time import sleep
# Ignore SSL certificate errors
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

#to create and write to CSV file
filename = "mtg_cards.csv"
f = open(filename, "w")

page_num = 1
serviceurl = "https://api.scryfall.com/cards/search?&page=%s\&q=" % page_num
card_type = "instant"#input("Enter card type here. All lower case. ")
cmdr = 'U'#input("Enter commander's colors, WUBRG format. ")

url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
print("Retrieving", url)
uh = urllib.request.urlopen(url, context=ctx)
data = uh.read().decode()
js = json.loads(data)
has_more = js["has_more"]
loop_count = 0
while has_more is True:
    count = 0
    for cards in range(175):
        name = js["data"][count]["name"]
        count += 1
        f.write(name+"\n")
        sleep(0.075)
    page_num += 1
    serviceurl = "https://api.scryfall.com/cards/search?&page=%s\&q=" % page_num
    url = serviceurl + "type%%3A%s" % card_type + "+" + "id%%3A%s" % cmdr
    uh = urllib.request.urlopen(url, context=ctx)
    data = uh.read().decode()
    js = json.loads(data)
    has_more = js["has_more"]
    loop_count += 1
else:
    try:
        count = 0
        for cards in range(js["total_cards"]-175):
            name = js["data"][count]["name"]
            count += 1
            f.write(name+"\n")
            sleep(0.075)
    except:
        pass

f.close()
