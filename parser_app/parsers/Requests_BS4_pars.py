from load_django import *

import re
import requests

from bs4 import BeautifulSoup
from openpyxl import Workbook

from parser_app.models import Product


headers = {"User-Agent": "Mozilla/5.0"}
url = "https://brain.com.ua/ukr/Mobilniy_telefon_Apple_iPhone_16_Pro_Max_256GB_Black_Titanium-p1145443.html"
# url = "https://brain.com.ua/ukr/Mobilniy_telefon_Xiaomi_Redmi_Note_14_8_256GB_Midnight_Black_1123261-p1190652.html" #for tests
# url = "https://brain.com.ua/ukr/Mobilniy_telefon_Xiaomi_Redmi_Note_14_8_256GB_Mist_Purple_1123263-p1190654.html" #for tests


response = requests.get(url)

soup = BeautifulSoup(response.text, "lxml")

product = {}

wb = Workbook()
ws = wb.active
ws.title = "Product"

try:
    product["Full name"] = (
        soup.find(name="span", string="Модель").find_next_sibling("span").text.strip()
    )
except AttributeError:
    product["Full name"] = None

try:
    product["Color"] = soup.find("a", title=lambda t: t and "Колір" in t).text.strip()
except AttributeError:
    product["Color"] = None

try:
    product["Memory size"] = soup.find(
        "a", title=lambda t: t and "Вбудована пам'ять" in t
    ).text.strip()
except AttributeError:
    product["Memory sizes"] = None

try:
    product["Price regular"] = " ".join(
        soup.find("div", class_="br-pr-price main-price-block").text.split()
    )
except AttributeError:
    product["Price regular"] = None

try:
    product["Price promo"] = soup.find("span", class_="red-price").text.strip()
except AttributeError:
    product["Price promo"] = None

try:
    img_links = soup.select("div.product-block-right img.br-main-img")
    product["Photo"] = [img["src"] for img in img_links if "src" in img.attrs]
except AttributeError:
    product["Photo"] = None

try:
    product["id"] = soup.find("span", class_="br-pr-code-val").text
except AttributeError:
    product["id"] = None

try:
    reviews = "".join(
        soup.find("a", class_="scroll-to-element brackets-reviews").text.strip()
    )
    product["Number of reviews"] = re.search(r"\d+", reviews).group()
except AttributeError:
    product["Number of reviews"] = None

try:
    product["Screen diagonal"] = soup.find(
        "a", title=lambda t: t and "Діагональ екрану" in t
    ).text.strip()
except AttributeError:
    product["Screen diagonal"] = None

try:
    product["Display resolution"] = soup.find(
        "a", title=lambda t: t and "Роздільна здатність екрану" in t
    ).text.strip()
except AttributeError:
    product["Display resolution"] = None


def clean(text):
    text = " ".join(text.stripped_strings)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s*,\s*", ", ", text)
    return text.strip()


try:
    specs = {}
    for block in soup.select(".br-pr-chr-item"):
        section = block.h3.get_text(strip=True)
        rows = {}
        for row in block.select("div > div"):
            spans = row.find_all("span")
            if len(spans) >= 2:
                key, value = clean(spans[0]), clean(spans[1])
                if "," in value:
                    value = [v.strip() for v in value.split(",")]
                rows[key] = value
        if rows:
            specs[section] = rows

    product["Characteristics"] = {"Characteristics": specs}
except AttributeError:
    product["Characteristics"] = None


print(product)


product_data = {
    "name": product.get("Full name"),
    "price_regular": product.get("Price regular"),
    "price_promo": product.get("Price promo"),
    "color": product.get("Color"),
    "memory": product.get("Memory size"),
    "manufacturer": product.get("Manufacturer"),
    "screen_diagonal": product.get("Screen diagonal"),
    "resolution": product.get("Display resolution"),
    "photos": product.get("Photo"),
    "characteristics": product.get("Characteristics"),
    "status": "Done",
}

obj, created = Product.objects.update_or_create(
    product_id=product.get("id"), defaults=product_data
)

if created:
    print(f"New write: {obj.name}")
else:
    print(f" Refresh: {obj.name}")
