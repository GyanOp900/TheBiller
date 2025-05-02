import os, json

def load_products():
    if not os.path.exists("products.json"):
        return []
    with open('products.json', 'r') as file:
        try:
            return json.load(file)
        except json.JSONDecodeError:
            return []

def save_products(products_list):
    with open("products.json", 'w') as file:
        json.dump(products_list, file, indent=4)