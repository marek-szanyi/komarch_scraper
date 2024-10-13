import requests
import re
from openpyxl import Workbook
from bs4 import BeautifulSoup
from lxml import etree


def parse_architect_details(detail_html_page):
    soup = BeautifulSoup(detail_html_page.content, "html.parser")
    kontakt_div = soup.find('div', string="Kontakt:")
    phone_number = ""
    if kontakt_div:
        phone_parent = kontakt_div.find_parent()
        phone_pattern = re.compile(r'''
               (.[+]*[(]{0,1}[0-9]{1,4}[)]{0,1}[-\s\./0-9]*)
           ''', re.VERBOSE)
        found_numbers = phone_pattern.findall(phone_parent.text)
        for anumber in found_numbers:
            phone_number += str(anumber).split('\n')[0].lstrip() + " "


    address_div = soup.find('div', string="Adresa:")
    address = None
    if address_div:
        address = address_div.find_next_sibling().text.strip()

    # Extract e-mail
    # //*[@id="app"]/main/div/div[2]/div[1]/div[3]/a
    dom = etree.HTML(str(soup))
    a_path = dom.xpath('//*[@id="app"]/main/div/div[2]/div[1]/div[3]/a')
    email = None
    if len(a_path) != 0:
        a_element = a_path[0]
        mailto = a_element.attrib['url']
        email = str(mailto).replace("mailto:","")

    return phone_number, address, email


# Example architect json object
# "id": 3045,
# "first_name": "Miloslav",
# "last_name": "Abel",
# "works_count": 0,
# "awards_count": 0,
# "contests_count": 0,
# "number": "0144 HA",
# "location_city": "Praha 5",
# "full_name": "Ing. arch. Miloslav Abel",
# "url": "https://www.komarch.sk/architekt/3045-ing-arch-miloslav-abel"
def parse_list_of_architects(sheet, json_data):
    for architect in json_data:
        phone_number = ""
        address = ""
        email = ""
        if architect.get("number") != "Zosnulý":
            print(architect.get("url"))
            (phone_number, address, email) = parse_architect_details(requests.get(architect.get("url")))

        # Add the necessary data extraction and writing logic here
        # Example: writing architect's name to a cell
        sheet.append([architect.get("id"), architect.get("number"), architect.get("full_name"), 
                      architect.get("location_city"), address, phone_number, email, architect.get("works_count"), 
                      architect.get("awards_count"), architect.get("contests_count"), architect.get("url")])

def main():
    # Create a new Excel document
    wb = Workbook()
    sheet = wb.active
    sheet.append(
        ["Id", "Registračné číslo", "Meno", "Miesto pôsobenia", "Adresa", "Telefónne číslo", "Email", "Vlastné diela",
         "Ocenenia", "Súťaže", "URL"])

    response = requests.get("https://www.komarch.sk/api/architects")
    response_json = response.json()
    next_link = response_json.get("links", {}).get("next")
    parse_list_of_architects(sheet, response_json.get("data"))
    while next_link:
        response = requests.get(next_link)
        response_json = response.json()
        parse_list_of_architects(sheet, response_json.get("data"))
        next_link = response_json.get("links", {}).get("next")

    # Save the document
    wb.save("architekti.xlsx")


if __name__ == "__main__":
    main()
