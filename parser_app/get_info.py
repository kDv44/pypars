from load_django import *
from parser_app.models import Product  # Ваша модель


def save_item(data):
    defaults = {
        "name": data["name"],
        "price": data["price"],
        "color": data["color"],
    }

    obj, created = Product.objects.get_or_create(
        product_id=data["id"],
        defaults=defaults,
    )

    obj.status = "Done"
    obj.save()
