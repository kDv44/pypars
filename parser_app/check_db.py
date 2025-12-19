from load_django import *
from parser_app.models import Product


def check_data():

    items = Product.objects.filter(status="Done").order_by("id")

    for item in items:
        print(f"ID: {item.product_id} | Name: {item.name} | Status: {item.status}")


if __name__ == "__main__":
    check_data()
