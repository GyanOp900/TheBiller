from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QGridLayout, QTreeWidget,
    QTreeWidgetItem, QScrollArea, QLineEdit,
    QFileDialog, QMessageBox, QMenuBar,
    QMenu
    

)
from PyQt6.QtGui import QPixmap, QFont, QIcon, QAction
from PyQt6.QtCore import Qt, QSize 
import sys                                 
from docxtpl import DocxTemplate
import datetime
import os, shutil
import subprocess
import platform
import json, qrcode
from pathlib import Path

import qrcode.constants
#import hashlib


doc = DocxTemplate("invoice_template.docx")

'''products = [
    {"name": "Pen", "price": 10, "image": "images/pen.png"},
    {"name": "Notebook", "price": 50, "image": "images/notebook.png"},
    {"name": "Eraser", "price": 5, "image": "images/eraser.png"},
    {"name": "Pencil", "price": 7, "image": "images/pencil.png"},
    {"name": "Ruler", "price": 15, "image": "images/ruler.png"},
    {"name": "Sharpener", "price": 8, "image": "images/sharpener.png"},
    {"name": "Marker", "price": 20, "image": "images/marker.png"},
    {"name": "Glue Stick", "price": 25, "image": "images/gluestick.png"},
    
    {"name": "Chocolate", "price": 30, "image": "images/chocolate.png"},
    {"name": "Chips", "price": 20, "image": "images/chips.png"},
    {"name": "Cookies", "price": 35, "image": "images/cookies.png"},
    {"name": "Juice Box", "price": 18, "image": "images/juicebox.png"},
    
    {"name": "Soap", "price": 28, "image": "images/soap.png"},
    {"name": "Toothpaste", "price": 45, "image": "images/toothpaste.png"},
    {"name": "Shampoo", "price": 90, "image": "images/shampoo.png"},
    {"name": "Hand Sanitizer", "price": 60, "image": "images/sanitizer.png"},
    
    {"name": "Earphones", "price": 299, "image": "images/earphones.png"},
    {"name": "USB Cable", "price": 120, "image": "images/usb.png"},
    {"name": "Power Bank", "price": 999, "image": "images/powerbank.png"},
    {"name": "Mouse", "price": 450, "image": "images/mouse.png"},
    
    {"name": "Mineral Water", "price": 20, "image": "images/water.png"},
    {"name": "Soda", "price": 25, "image": "images/soda.png"},
    {"name": "Energy Drink", "price": 80, "image": "images/energydrink.png"},
]'''  # this is how the products are saved in products.json just a bit cleaner.

with open("products.json", 'r') as file:
    products = json.load(file)


global LIGHT_STYLE, DARK_STYLE

DARK_STYLE = """
QWidget {
    background-color: #1b1b27;
    color: #E0E0E0;
    font-family: 'Segoe UI', 'Roboto', sans-serif;
    font-size: 14px;
}

QLabel {
    font-size: 14px;
}

QPushButton {
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2e2e45, stop:1 #3e3e5e);
    color: white;
    border-radius: 12px;
    padding: 8px 16px;
    border: 1px solid #4c4c6f;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #5c5c8a;
    border: 1px solid #6e6eb0;
}

QLineEdit, QTreeWidget, QTextEdit {
    background-color: #2a2a3f;
    border: 1px solid #555575;
    border-radius: 10px;
    padding: 6px;
    color: white;
}

QHeaderView::section {
    background-color: #33334c;
    color: white;
    padding: 6px;
    border: 1px solid #555575;
    border-radius: 5px;
}

QScrollBar:vertical {
    background: #1b1b27;
    width: 10px;
    margin: 0;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #5a5a8a, stop:1 #7070aa);
    border-radius: 5px;
}

QScrollBar::handle:vertical:hover {
    background: #8a8ac5;
}
"""



LIGHT_STYLE = """
QWidget {
    background-color: #f7f9fc;
    color: #2c2c2c;
    font-family: 'Segoe UI', 'Roboto', sans-serif;
    font-size: 14px;
}

QLabel {
    font-size: 14px;
}

QPushButton {
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ffffff, stop:1 #e4e6ef);
    border: 1px solid #d0d0d0;
    border-radius: 12px;
    padding: 8px 16px;
    font-weight: bold;
    color: #333;
}

QPushButton:hover {
    background-color: #dfe3ec;
    border: 1px solid #b0b0b0;
}

QLineEdit, QTreeWidget, QTextEdit {
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 10px;
    padding: 6px;
    color: #222;
}

QHeaderView::section {
    background-color: #f0f0f0;
    color: black;
    padding: 6px;
    border: 1px solid #ccc;
    border-radius: 5px;
}

QScrollBar:vertical {
    background: #f4f4f4;
    width: 10px;
    margin: 0;
    border-radius: 5px;
}

QScrollBar::handle:vertical {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #bbbbbb, stop:1 #999999);
    border-radius: 5px;
}

QScrollBar::handle:vertical:hover {
    background: #888888;
}
"""




class BillingSystem(QMainWindow):
    def __init__(self, app):
        super().__init__()
        
        self.setWindowTitle("Billing System (V0.8 ~THE QR UPDATE~)")
        self.app = app
        self.setGeometry(0, 0, 900, 600)
        self.showMaximized()
        self.cart = {}
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)

        self.main_layout = QHBoxLayout()
        self.main_widget.setLayout(self.main_layout)

        self.create_menu_bar()

        self.create_cart_area()
        
        self.create_product_area()
        
        try:
            with open('settings.json', 'r') as f:
                theme_settings = json.load(f)
            if theme_settings.get('dark_mode', False):
                self.app.setStyleSheet(DARK_STYLE)
            else:
                self.app.setStyleSheet(LIGHT_STYLE)
        except FileNotFoundError:
            pass




    def create_menu_bar(self):
        '''menu_bar = self.menuBar()
        menu_bar.setStyleSheet("""
                                background-color: white;
                                color: black;""")

        file_menu = menu_bar.addMenu("File")
        #settings_menu = menu_bar.addMenu("Settings")

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)'''
        pass

    def create_product_area(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        grid = QGridLayout()
        content_widget.setLayout(grid)



        #global product
        for index, product in enumerate(products):
            row = index // 2
            col = index % 2

            vbox = QVBoxLayout()
            name_label = QLabel(product['name'])
            price_label = QLabel(f"Price : ₹{product['price']}")
            stock_label = QLabel(f"Stock : {product['stock']}")
            product['stock_label'] = stock_label

            pixmap = QPixmap(product['image']).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            img_label = QLabel()
            img_label.setPixmap(pixmap)
            img_label.setStyleSheet("border: 1px solid gray; padding : 5px;")

            btn = QPushButton("Add to Cart")
            btn.setStyleSheet("""
                                QPushButton {
                            padding: 5px;
                            font: 19px;
                              }""")
            btn.clicked.connect(lambda checked, p=product: self.add_to_cart(p))

            
            

            vbox.addWidget(img_label)
            vbox.addWidget(name_label)
            vbox.addWidget(price_label)
            vbox.addWidget(stock_label)
            vbox.addWidget(btn)

            grid.addLayout(vbox, row, col)

        scroll.setWidget(content_widget)
        
        self.main_layout.addWidget(scroll, 2)

    def create_cart_area(self):
        cart_layout = QVBoxLayout()
        cart_label = QLabel("C  A  R  T")
        cart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cart_label.setStyleSheet("font: bold 17px")

        self.tree = QTreeWidget()
        self.tree.setStyleSheet("font: bold 14px")
        self.tree.setHeaderLabels(['Item', 'Price', 'Qty', 'Total'])
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.MultiSelection)

        font = QFont("Arial", 12)
        


        self.tree.setFont(font)


        self.pay_btn = QPushButton("Pay")
        self.pay_btn.setStyleSheet("""
                                    QPushButton {
                                   background-color: #4e9eff;
                                   border: 5px solid #262626;
                                   border-radius: 20px;
                                   padding: 10px;
                                   font: bold 15px;
                                   color: #ffffff;
                                   }
                                   
                                   QPushButton:hover {
                                   background-color: #3b8df2;
                                   border: 2px solid #262626;
                                   border-radius: 20px;
                                   padding: 10px;
                                   font: bold 15px;
                                   color: #ffffff;
                                   }""")
        self.pay_btn.clicked.connect(self.show_qr_code)

        self.delete_selected = QPushButton("Delete Selected Item")
        self.delete_selected.clicked.connect(self.deletesel)
        self.delete_selected.setStyleSheet("""
                                    QPushButton {
                                   background-color: #4e9eff;
                                   border: 5px solid #262626;
                                   border-radius: 20px;
                                   padding: 10px;
                                   font: bold 15px;
                                   color: #ffffff;
                                   }
                                   
                                   QPushButton:hover {
                                   background-color: #3b8df2;
                                   border: 2px solid #262626;
                                   border-radius: 20px;
                                   padding: 10px;
                                   font: bold 15px;
                                   color: #ffffff;
                                   }""")

        self.info_label = QLabel("The store's helper..")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.info_label.setStyleSheet("""
                                        QLabel {
                                    font: bold 14px;  
                                      }""")
        
        self.settings_btn = QPushButton()
        self.settings_btn.clicked.connect(self.open_settings)
        self.settings_btn.setIcon(QIcon('images/setting_icon.png'))
        self.settings_btn.setIconSize(QSize(38,38))
        self.settings_btn.setStyleSheet("""
                                        OPushButton {
                                        border-radius: 20px;
                                        }""")
        self.settings_btn.setFixedSize(70,50)

        cart_layout.addWidget(self.settings_btn)
        cart_layout.addWidget(cart_label)
        cart_layout.addWidget(self.tree)
        cart_layout.addWidget(self.pay_btn)
        cart_layout.addWidget(self.delete_selected)
        cart_layout.addWidget(self.info_label)

        cart_widget = QWidget()
        cart_widget.setLayout(cart_layout)

        self.main_layout.addWidget(cart_widget, 1)

    def show_qr_code(self):
        amount = sum(price * qty for price, qty in self.cart.values())

        with open('settings.json' , 'r') as f:
            settings = json.load(f)
            upi_id = settings['upi_idd']
            name = settings['name']

        if not upi_id:
            QMessageBox.warning(self, "Input Error", "Please enter an upi id!")
        

        # this two 'self.' had me starving for 3 hours MAN THIS IS SO SHITT THIS WAS THE WHOLE FRICKING PROBLEM
        self.qr_window = ShowQR(upi_id=upi_id, amount=amount, name=name, billing_system=self, app=self.app, )
        self.qr_window.show()

    def add_to_cart(self, product):
        name = product['name']
        price = product['price']

        if product['stock'] <= 0:
            QMessageBox.warning(self, "No Stock!", f"That product is not available! Stock {product['stock']}")
            return
        else:
            product['stock'] -= 1
            self.update_stock(product=product)

            stock_label = product.get('stock_label')

            if stock_label:
                stock_label.setText(f"Stock : {product['stock']}")

            if name in self.cart:
                self.cart[name][1] += 1
            else:
                self.cart[name] = [price, 1]
            self.update_cart()

    def update_cart(self):
        self.tree.clear()
        for name, (price, qty) in self.cart.items():
            total = price * qty
            QTreeWidgetItem(self.tree, [name, f"₹{price}", str(qty), f"₹{total}"])
    
    def update_stock(self, product):
        from utils import load_products, save_products
        productsss = load_products()

        for item in productsss:
            if item['name'].lower() == product['name'].lower():
                item['stock'] = product['stock']
        
        save_products(productsss)




    def show_invoice(self):
        total = sum(price * qty for price, qty in self.cart.values())
        invoice_text = "\n".join([
            f"{name} x{qty} = ₹{price * qty}" for name, (price, qty) in self.cart.items()
        ])
        invoice_text += f"\n\nTotal: ₹{total}"
        print("\n===== Invoice =====")
        print(invoice_text)

        invoice_list = []

        for name, (price, qty) in self.cart.items():
            line_total = price * qty
            invoice_list.append([qty, name, price, line_total])

        with open('settings.json' , 'r') as f:
            settings = json.load(f)

        doc.render({
                    "shopname" : settings['shopname'],
                    "address" : settings['address'],
                    "name" : settings['name'], 
                    "phone" : settings['phone'],
                    "invoice_list" : invoice_list,
                    "subtotal" :total,
                    "salestax" :"18%",
                    "total" : total + (total * 0.1),
                    "datenow" : datetime.datetime.now().strftime('%y-%m-%d'),
                    "timenow" : datetime.datetime.now().strftime('%H:%M')
                    })
        
        save_folder = "invoices"
        os.makedirs(save_folder, exist_ok=True)

        filename = f"{datetime.datetime.now().strftime('%y-%m-%d_%H-%M-%S')}" + "invoice.docx"

        filepath = os.path.join(save_folder, filename)
        doc.save(filepath)
        

    
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":  
            subprocess.call(["open", filepath])
        else:  
            subprocess.call(["xdg-open", filepath])

    def deletesel(self):
        selected = self.tree.selectedItems()

        if not selected:
            return
        
        for item in selected:
            name = item.text(0)
            if name in self.cart:
                del self.cart[name]

        self.update_cart()

    def open_settings(self):
        self.window = SettingsWindow(self.app)
        self.window.show()

    def add_product(self):
        Product_Window = ProductWindow(self.app)
        Product_Window.show()

class ProductWindow(QWidget):
    def __init__(self,app):
        super().__init__()
        self.setGeometry(100,100,300,300)
        self.setWindowTitle("Add Product")
        self.setWindowIcon(QIcon('setting_icon.ico'))

        layout = QVBoxLayout()

        title_label = QLabel("Add Product")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("""padding: 10px; font: bold 15px;""")
        title_label.setMaximumHeight(30)

        self.product_name_entry = QLineEdit()
        self.product_name_entry.setPlaceholderText("Enter Product name here")

        self.product_price_entry = QLineEdit()
        self.product_price_entry.setPlaceholderText("Enter Product price here")

        self.product_stock_entry = QLineEdit()
        self.product_stock_entry.setPlaceholderText("Enter Product stock here")

        self.image_upload_widget = ImageUploadWidget()

        self.save_button = QPushButton("Add Product")
        self.save_button.clicked.connect(self.save_product)

        self.restart_label = QLabel("Restart Required")
        self.restart_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.restart_label.setStyleSheet("""
                                        padding: 10px;
                                        font: bold 10px""")





        layout.addWidget(title_label)
        layout.addWidget(self.product_name_entry)
        layout.addWidget(self.product_price_entry)
        layout.addWidget(self.product_stock_entry)
        layout.addWidget(self.image_upload_widget)
        layout.addWidget(self.save_button)
        layout.addWidget(self.restart_label)
        self.setLayout(layout)

    def load_products(self):
        if not os.path.exists("products.json"):
            return []
        with open('products.json', 'r') as file:
            try:
                return json.load(file)
            except json.JSONDecodeError:
                return []
    
    def save_products(self, products_list):
        with open("products.json", 'w') as file:
            json.dump(products_list, file, indent=4)



    def save_product(self):
        Pname = self.product_name_entry.text().strip()
        Pprice = self.product_price_entry.text().strip()
        Pimage = self.image_upload_widget.return_img_path()
        Pstock = self.product_stock_entry.text().strip()




        if not Pname or not Pprice or not Pimage or not Pstock:
            QMessageBox.warning(self, "Input Error", "Please enter all the fields above!")
            return
        
        try:
            Pprice = float(Pprice)
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Please enter a valid number!")
            return
        
        try:
            Pstock = int(Pstock)
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Please enter a valid stock!")
            return
        


        new_product = {
                      "name" : Pname,
                      "price" : Pprice,
                      "image" : Pimage,
                      "stock" : Pstock
                      }
        
        products = self.load_products()
        products.append(new_product)
        self.save_products(products)
        QMessageBox.warning(self, "Sucessful", "Product Was added!")


        

class ImageUploadWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):

        layout = QVBoxLayout()

        self.upload_button = QPushButton("Upload Image")
        self.upload_button.setStyleSheet("""
                                        QPushButton {
                                         border: 2px dashed #aaa;
                                         padding: 40px;
                                         font-size: 16px;
                                         }""")
        

        self.upload_button.clicked.connect(self.open_file_dialog)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.upload_button)
        layout.addWidget(self.image_label)

        self.setLayout(layout)

    def open_file_dialog(self):
        options = QFileDialog.Option.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Image",
            "",
            "PNG Files (*.png);;All Files (*)",
            options=options
        )

        if file_path:
            self.save_image(file_path)
    def save_image(self, source_path):
        images_dir = Path("images")
        images_dir.mkdir(exist_ok=True)
        
        file_name = Path(source_path).name
        destination_path = images_dir / file_name
        
        self.image_path = str(Path(destination_path)).replace("\\", "/")

        if Path(source_path).resolve() != destination_path.resolve():
            shutil.copy(source_path, destination_path)

        self.display_image(destination_path)
        

    def display_image(self, image_path):

        image_path = str(image_path)

        if not image_path or not os.path.exists(image_path):
            QMessageBox.warning(self, "Image Error", "Image file not found or invalid!")
            return

        pixmap = QPixmap(image_path)
        if pixmap.isNull():
            QMessageBox.warning(self, "Image Error", "Failed to load the image.")
            return

        self.image_label.setPixmap(pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        self.upload_button.hide()

    def return_img_path(self):
        try:
            return self.image_path
        except AttributeError:
            QMessageBox.warning(self, "Input Error", "Upload an image!")
            return ""

class DeleteProductWindow(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.setGeometry(200,200,200,100)
        self.setWindowTitle("Delete Product")
        self.setWindowIcon(QIcon("setting_icon.ico"))

        layout = QVBoxLayout()
        name_label = QLabel(" Enter Product Name ")
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setStyleSheet("""padding: 5px; font: bold 19px;""")

        self.product_name = QLineEdit()

        self.delete_btn = QPushButton("Delete Product")
        self.delete_btn.clicked.connect(self.delete_product)

        layout.addWidget(name_label)
        layout.addWidget(self.product_name)
        layout.addWidget(self.delete_btn)
        self.setLayout(layout)
    
    def delete_product(self):
        from utils import load_products, save_products
        product_name = self.product_name.text()

        if not product_name:
            QMessageBox.warning(self,"Input Error", "Please enter a product name!")

        products = load_products()

        found = False
        for item in products:
            if item['name'].lower() == product_name.lower():
                products.remove(item)
                found = True
                break
        
        if found:
            QMessageBox.warning(self, "Product Removed", f"{product_name} was removed! Please restart for changes to be applied!")
        else:
            QMessageBox.warning(self, "Not found", f"{product_name} was not found!")


        save_products(products)

class StockWindow(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.setWindowTitle("Add Stock Window")
        self.setGeometry(100,100,300,200)
        self.setWindowIcon(QIcon("setting_icon.ico"))

        layout = QVBoxLayout()


        self.product_name = QLineEdit()
        self.product_name.setPlaceholderText("Enter Product Name")


        self.product_stock = QLineEdit()
        self.product_stock.setPlaceholderText("Enter Product Stock")

        self.save_btn_stock = QPushButton("Save")
        self.save_btn_stock.clicked.connect(self.save_stock)



        layout.addWidget(self.product_name)
        layout.addWidget(self.product_stock)
        layout.addWidget(self.save_btn_stock)
        self.setLayout(layout)

    def save_stock(self):
        name = self.product_name.text()
        stock = self.product_stock.text()

        try:
            stock = int(stock)
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid Input")

        from utils import load_products, save_products

        productss = load_products()
        founds = False

        #stock_label = product.get("stock_label")
        
        #stock_label.setText(f"Stock : {stock}")
        
        
        for item in productss:
            if item['name'].lower() == name.lower():
                item['stock'] += stock
                founds = True
                QMessageBox.warning(self,"Done","Saved Sucessfully")
                break

        save_products(productss)



        


class SettingsWindow(QWidget):

    


    def __init__(self, app):
        super().__init__()
        self.app = app
        self.setWindowTitle("Settings")
        self.setGeometry(200,200,300,200)
        self.setWindowIcon(QIcon('setting_icon.ico'))
        #self.setFixedSize(300,400)

        layout = QVBoxLayout()

        name_label = QLabel("Name: ")
        self.name_input = QLineEdit()

        phone_label = QLabel("Phone: ")
        self.phone_input = QLineEdit()

        shopname_label = QLabel("Shop Name: ")
        self.shop_input = QLineEdit()

        self.upi_id = QLineEdit()
        self.upi_id.setPlaceholderText("Enter UPI ID HERE")

        address_label = QLabel("Address: ")
        self.address_input = QLineEdit()


        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_settings)

        line_label = QLabel("_________________________________________________________")
        line_label1 = QLabel("_________________________________________________________")
        line_label2 = QLabel("_________________________________________________________")
        


        self.dark_mode = False

        self.theme_button = QPushButton("Switch To Dark Mode")
        self.theme_button.clicked.connect(self.toggle_theme)

        feature_warning_label = QLabel("This feature may have some bugs!")
        feature_warning_label.setStyleSheet("""
                                    QLabel {
                                    font: bold 15px
                                    }""")
        
        self.product_btn = QPushButton("Add Product")
        self.product_btn.clicked.connect(self.open_product_window)

        self.del_product_btn = QPushButton("Delete a Product")
        self.del_product_btn.clicked.connect(self.open_del_product_window)

        self.add_stock_btn = QPushButton("Add Stock For an Item")
        self.add_stock_btn.clicked.connect(self.open_add_stock_window)

        

        layout.addWidget(name_label)
        layout.addWidget(self.name_input)
        #layout.addWidget(line_label)
        layout.addWidget(phone_label)
        layout.addWidget(self.phone_input)
        layout.addWidget(shopname_label)
        layout.addWidget(self.shop_input)
        layout.addWidget(address_label)
        layout.addWidget(self.address_input)
        layout.addWidget(self.upi_id)
        layout.addWidget(save_button)
        layout.addWidget(line_label1)
        layout.addWidget(self.theme_button)
        layout.addWidget(feature_warning_label)
        layout.addWidget(self.product_btn)
        layout.addWidget(line_label2)
        layout.addWidget(self.del_product_btn)
        layout.addWidget(self.add_stock_btn)


        self.setLayout(layout)

    def open_add_stock_window(self):
        self.swindow = StockWindow(self.app)
        self.swindow.show()

    def open_del_product_window(self):
        self.dwindow = DeleteProductWindow(self.app)
        self.dwindow.show()



    def open_product_window(self):
        self.pwindow = ProductWindow(self.app)
        self.pwindow.show()

    def save_settings(self):
        
        settings = {
            'name' : self.name_input.text(),
            'phone' : self.phone_input.text(),
            'shopname' : self.shop_input.text(),
            'address' : self.address_input.text(),
            'upi_idd' : self.upi_id.text()
                    }

        with open('settings.json', 'w') as f:
            json.dump(settings, f, indent=4)

        self.close()

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode # If dark mode is true is will set to dark mode else light mode
        self.change_theme(self.dark_mode)

        try:
            with open('settings.json', 'r') as f:
                settings = json.load(f)
        except FileNotFoundError:
            settings = {}

        settings['dark_mode'] = self.dark_mode


        with open('settings.json', 'w') as f:
            json.dump(settings, f, indent=4)

    


    def change_theme(self, mode):
        if mode:
            self.app.setStyleSheet(DARK_STYLE)
            self.theme_button.setText("Switch to Light mode")
        else:
            self.app.setStyleSheet(LIGHT_STYLE)
            self.theme_button.setText("Switch to Dark mode")

class ShowQR(QWidget):
    def __init__(self, upi_id, amount, name, app, billing_system, currency='INR'):
        super().__init__()
        self.app = app
        self.setGeometry(100,100,300,200)
        self.setWindowTitle("Payment Window")
        self.upi_id = upi_id
        self.amount = amount
        self.currency = currency
        self.name = name
        self.billing_system = billing_system

        qr_path = self.make_qr()

        layout = QVBoxLayout()

        Pay_here_label = QLabel("Scan and Pay")
        Pay_here_label.setStyleSheet("""font: bold 14px""")
        Pay_here_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.qr_label = QLabel()
        pixmap = QPixmap(qr_path)
        if pixmap.isNull():
            QMessageBox.warning(self, "QR Error", "Could not load QR code image.")
        else:
            self.qr_label.setPixmap(pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        #.scaled(100,100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.qr_label.setPixmap(pixmap)

        self.generate_reciept_btn = QPushButton("Generate Reciept")
        self.generate_reciept_btn.setStyleSheet("""
                                                font: bold 14px;
                                                padding: 10px;
                                                """)
        self.generate_reciept_btn.clicked.connect(self.payment_done)
        

        layout.addWidget(Pay_here_label)
        layout.addWidget(self.qr_label)
        layout.addWidget(self.generate_reciept_btn)
        self.setLayout(layout)


    def payment_done(self):
        
        reply = QMessageBox.question(self, "Confirm Payment", "Has the payment been completed?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.billing_system.show_invoice()
            self.close()

    def make_qr(self):
        main_data = f"upi://pay?pa={self.upi_id}&pn={self.name}&am={self.amount}&cu={self.currency}"

        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            border=4,
            box_size=10,
        )
        
        qr.add_data(main_data)

        qr.make(fit=True)

        img = qr.make_image(fill_color = 'black', back_color = 'white')   


        path = "upi_payment.png" 


        img.save(path)

        return path





    




from login import LoginWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)

    login = LoginWindow(app=app)
    login.show()
    sys.exit(app.exec())
