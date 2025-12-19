from load_django import *

import re
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

from parser_app.models import Product


options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

driver.get("https://brain.com.ua/ukr/")

product = {}

time.sleep(2)

inputs = driver.find_elements(By.CLASS_NAME, "quick-search-input")
input_search = None

for inp in inputs:
    if inp.is_displayed():
        input_search = inp
        break

if input_search:
    input_search.click()
    input_search.clear()
    input_search.send_keys("Apple iPhone 15 128GB Black")
    input_search.send_keys(Keys.ENTER)
else:
    print("Не найдено видимое поле поиска!")
    driver.quit()
    exit()


time.sleep(3)

first_product = driver.find_element(
    By.XPATH, "(//div[contains(@class, 'product-wrapper')]//a)[1]"
)
first_product.click()

time.sleep(3)


try:
    product["Full name"] = (
        driver.find_element(By.TAG_NAME, "h1").get_attribute("textContent").strip()
    )
except:
    product["Full name"] = None

try:
    product["Color"] = (
        driver.find_element(By.XPATH, "//span[contains(text(), 'Колір')]/../span[2]")
        .get_attribute("textContent")
        .strip()
    )
except:
    product["Color"] = None

try:
    product["Memory size"] = (
        driver.find_element(
            By.XPATH, '//span[contains(text(), "Вбудована пам\'ять")]/../span[2]'
        )
        .get_attribute("textContent")
        .strip()
    )
except:
    product["Memory size"] = None

try:
    product["Manufacturer"] = (
        driver.find_element(By.XPATH, "//span[contains(text(), 'Виробник')]/../span[2]")
        .get_attribute("textContent")
        .strip()
    )
except:
    product["Manufacturer"] = None

try:
    price_container = driver.find_element(By.CLASS_NAME, "br-pr-price")
    try:
        promo = driver.find_element(By.CLASS_NAME, "red-price")
        product["Price promo"] = promo.text.replace(" ", "").replace("грн", "").strip()
        product["Price regular"] = (
            driver.find_element(By.CLASS_NAME, "old-price")
            .text.replace(" ", "")
            .replace("грн", "")
            .strip()
        )
    except:
        product["Price promo"] = None
        product["Price regular"] = (
            price_container.text.replace(" ", "")
            .replace("грн", "")
            .replace("\n", "")
            .strip()
        )
except:
    product["Price regular"] = None
    product["Price promo"] = None

try:
    img_elems = driver.find_elements(
        By.CSS_SELECTOR, ".br-pr-img-labels-container img, .br-main-img"
    )
    product["Photo"] = list(
        set([img.get_attribute("src") for img in img_elems if img.get_attribute("src")])
    )
except:
    product["Photo"] = []

try:
    product["id"] = (
        driver.find_element(By.CLASS_NAME, "br-pr-code-val")
        .get_attribute("textContent")
        .strip()
    )
except:
    product["id"] = None

try:
    rev_text = driver.find_element(By.CSS_SELECTOR, "a.reviews-count").get_attribute(
        "textContent"
    )
    product["Number of reviews"] = (
        re.search(r"\d+", rev_text).group() if re.search(r"\d+", rev_text) else "0"
    )
except:
    product["Number of reviews"] = "0"

try:
    product["Screen diagonal"] = (
        driver.find_element(
            By.XPATH, "//span[contains(text(), 'Діагональ екрану')]/../span[2]"
        )
        .get_attribute("textContent")
        .strip()
    )
except:
    product["Screen diagonal"] = None

try:
    product["Display resolution"] = (
        driver.find_element(
            By.XPATH,
            "//span[contains(text(), 'Роздільна здатність екрану')]/../span[2]",
        )
        .get_attribute("textContent")
        .strip()
    )
except:
    product["Display resolution"] = None

driver.execute_script("window.scrollBy(0, 600);")
time.sleep(1)

specs = {}
spec_blocks = driver.find_elements(By.XPATH, "//div[contains(@class,'br-pr-chr-item')]")

for block in spec_blocks:
    rows = block.find_elements(By.XPATH, ".//div[span]")
    for row in rows:
        spans = row.find_elements(By.TAG_NAME, "span")
        if len(spans) >= 2:
            key = spans[0].text.strip()
            value = spans[1].get_attribute("textContent").strip()

            links = spans[1].find_elements(By.TAG_NAME, "a")
            if links:
                value = ", ".join([a.text.strip() for a in links if a.text.strip()])

            if key:
                specs[key] = value

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
