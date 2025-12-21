from django.db import models

# Create your models here.
from django.db import models


class Product(models.Model):
    product_id = models.CharField(max_length=100, unique=True)
    name = models.CharField(max_length=255)
    price_regular = models.CharField(max_length=50, null=True, blank=True)
    price_promo = models.CharField(max_length=50, null=True, blank=True)
    color = models.CharField(max_length=100, null=True, blank=True)
    memory = models.CharField(max_length=100, null=True, blank=True)
    manufacturer = models.CharField(max_length=100, null=True, blank=True)
    screen_diagonal = models.CharField(max_length=100, null=True, blank=True)
    resolution = models.CharField(max_length=100, null=True, blank=True)
    number_of_reviews = models.CharField(max_length=100, null=True, blank=True)

    photos = models.JSONField(default=list, null=True, blank=True)
    characteristics = models.JSONField(default=dict, null=True, blank=True)

    url = models.URLField(max_length=500, null=True, blank=True)
    status = models.CharField(max_length=20, default="New")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name
