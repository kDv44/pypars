from load_django import *

import re
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

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
    raw_name = driver.find_element(By.TAG_NAME, "h1").get_attribute("textContent")
    product["Full name"] = (
        raw_name.replace("Мобільний телефон ", "").split("(")[0].strip()
    )
except Exception:
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
        re.search(r"\d+", rev_text).group() if re.search(r"\d+", rev_text) else None
    )
except:
    product["Number of reviews"] = None

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

try:
    product["URL"] = driver.find_element(
        By.XPATH, "//link[@rel='canonical']"
    ).get_attribute("href")
except:
    product["URL"] = None

driver.execute_script("window.scrollBy(0, 600);")
time.sleep(1)

driver.execute_script("window.scrollBy(0, 600);")
time.sleep(2)

try:

    wait = WebDriverWait(driver, 10)
    all_specs_button = wait.until(
        ec.presence_of_element_located(
            (
                By.XPATH,
                "//button[contains(@class, 'br-prs-button')][.//span[text()='Всі характеристики']]",
            )
        )
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", all_specs_button
    )
    time.sleep(1)

    driver.execute_script("arguments[0].click();", all_specs_button)

    print("Кнопку 'Всі характеристики' натиснуто")
    time.sleep(2)
except Exception as e:
    print(f"Кнопку не вдалося натиснути (можливо, список вже відкритий): {e}")


specs = {}


spec_blocks = driver.find_elements(By.XPATH, "//div[contains(@class,'br-pr-chr-item')]")

for block in spec_blocks:
    try:
        section_name = block.find_element(By.TAG_NAME, "h3").text.strip()
    except:
        section_name = "Загальне"

    section_data = {}
    rows = block.find_elements(By.XPATH, ".//div[span]")

    for row in rows:
        spans = row.find_elements(By.TAG_NAME, "span")
        if len(spans) >= 2:
            key = spans[0].text.strip()

            raw_value = spans[1].get_attribute("textContent")
            value = " ".join(raw_value.replace("\xa0", " ").split()).strip()

            links = spans[1].find_elements(By.TAG_NAME, "a")
            if links:
                link_texts = [a.text.strip() for a in links if a.text.strip()]
                if link_texts:
                    value = ", ".join(link_texts)

            if key:
                section_data[key] = value

    if section_data:
        specs[section_name] = section_data

product["Characteristics"] = specs


print(product)

product_data = {
    "name": product.get("Full name"),
    "color": product.get("Color"),
    "price_regular": product.get("Price regular"),
    "price_promo": product.get("Price promo"),
    "memory": product.get("Memory size"),
    "manufacturer": product.get("Manufacturer"),
    "photos": product.get("Photo"),
    "number_of_reviews": product.get("Number of reviews"),
    "screen_diagonal": product.get("Screen diagonal"),
    "resolution": product.get("Display resolution"),
    "characteristics": product.get("Characteristics"),
    "url": product.get("URL"),
    "status": "Done",
}

obj, created = Product.objects.update_or_create(
    product_id=f"SEL_{product.get("id")}", defaults=product_data
)

if created:
    print(f"New write: {obj.name}")
else:
    print(f" Refresh: {obj.name}")
