import re
import time

import requests
from openpyxl import Workbook
from openpyxl.styles import Font



GIST_URL = ("https://gist.github.com/captn3m0/4f3da8f07fe884e62bfab3ac85616936/raw/"
            "27f26717533ba451e56a8769fb0eb224605c12a5/01-domains.md")
SUFFIXES = ("gov.in", "nic.in")                 
HOST_PATTERN = re.compile(
    r"(?<![a-z0-9@.-])(?:[a-z0-9-]+\.)+(?:gov|nic)\.in(?![a-z0-9-]|\.[a-z0-9])", re.I)



def download():
    for attempt in range(1, 4):
        try:
            reply = requests.get(GIST_URL, timeout=60)
            reply.raise_for_status()
            return reply.text
        except requests.RequestException as error:
            print(f"try {attempt}/3 failed: {type(error).__name__}")
            time.sleep(5)
    raise SystemExit("Could not download the list. Check your internet connection.")



def save(name, urls, title):
    with open(name + ".txt", "w", encoding="utf-8") as file:
        file.write("\n".join(urls))

    book = Workbook()
    sheet = book.active
    sheet.title = title
    sheet.append(["S.No", "Website URL"])
    for number, url in enumerate(urls, start=1):
        sheet.append([number, url])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.column_dimensions["B"].width = 55
    try:
        book.save(name + ".xlsx")
    except PermissionError:                      
        print(f"Close {name}.xlsx in Excel and run again.")



text = download()

hosts = {h.lower() for h in HOST_PATTERN.findall(text)} 

if not hosts:                                    
    print("No names found. The first 300 characters downloaded were:")
    print(repr(text[:300]))
    raise SystemExit("Stopped, no files were written.")

for suffix in SUFFIXES:                          
    urls = sorted("https://" + h for h in hosts
                  if h == suffix or h.endswith("." + suffix))
    name = suffix.replace(".", "_") + "_urls"    
    if not urls:                                
        print(f"{suffix}: no URLs found, files not written.")
        continue
    save(name, urls, suffix)
    print(f"{suffix}: saved {len(urls)} URLs -> {name}.txt, {name}.xlsx")
