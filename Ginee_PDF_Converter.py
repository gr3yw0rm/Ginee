import os
import re
import io
import glob
import time
import json
import fitz
import qrcode
import barcode
from barcode.writer import ImageWriter
import pandas as pd
import numpy as np
import datetime as dt
from PIL import Image, ImageFont, ImageDraw
import textwrap
import webbrowser


# Folder locations
downloads_folder = os.path.join(os.environ.get('HOMEPATH'), 'Downloads')
onedrive_folder = os.path.join(os.getenv('HOMEPATH'), 'OneDrive', 'Shared Files - Shop', 'Python Scripts', 'Ginee')

# Thank you image file
with Image.open(os.path.join(onedrive_folder, 'ty for your purchase.jpg'), 'r') as ty_image:
    imgByteArr = io.BytesIO()
    # ty_image = ty_image.rotate(270)
    ty_image.save(imgByteArr, format='PNG')
    ty_image_bytes = imgByteArr.getvalue()

def draw_text(text, size, font_size, wraptext=None):
    if wraptext:
        lines = textwrap.wrap(text, width=wraptext)
        text = '\n'.join(lines)

    imgByteArr = io.BytesIO()
    font = ImageFont.truetype(os.path.join(onedrive_folder, 'LiberationSans-Regular.ttf'), size=font_size)
    text_image = Image.new(mode='RGB', size=size, color='#ffffff')
    draw = ImageDraw.Draw(im=text_image)
    draw.text(xy=(size[0]/2, size[1]/2), text=text, font=font, fill='black', anchor='mm', align='left')
    # text_image = text_image.rotate(270, expand=1)
    # text_image.show()
    text_image.save(imgByteArr, format='PNG')
    return imgByteArr.getvalue()


def draw_greetings(customer_name):
    formatted_customer_name = re.sub(' [a-zA-Z]?\.* ', ' ', customer_name).strip().split('/')[0]
    print(formatted_customer_name)
    splitted_customer_name = formatted_customer_name.split(' ')
    for i, name in enumerate(splitted_customer_name):
        joined_names = ' '.join(splitted_customer_name[:i])
        if len(joined_names) <= 12:
            greeting_name = joined_names
    imgByteArr = io.BytesIO()
    font = ImageFont.truetype(font=os.path.join(onedrive_folder, 'MarckScript-Regular.ttf'), size=72)   # customized font
    greeting_image = Image.new(mode='RGB', size=(640, 104), color='#ffffff')
    draw = ImageDraw.Draw(im=greeting_image)
    draw.text(xy=(320, 50), text=f"Hi {greeting_name.title()}!", font=font, fill='black', anchor='mm')
    # greeting_image.show()
    greeting_image.save(imgByteArr, format='PNG')
    return imgByteArr.getvalue()


def generate_barcode(text, type, barcode_type='Code39', write_text=True):
    """Generates barcode in bytes"""
    buffered = io.BytesIO()
    if type == 'barcode':
        barcode.base.Barcode.default_writer_options['write_text'] = write_text
        barcode.generate(barcode_type, text, writer=ImageWriter(), output=buffered)
    elif type == 'qrcode':
        qr = qrcode.QRCode(box_size=20)
        qr.add_data(text)
        img = qr.make_image(back_color = 'Transparent')
        img.save(buffered, format='PNG', transperancy=0, fill=(255, 0, 0))
    return buffered.getvalue()
        

def draw_order_details(order_details, size=(596, 300)):
    imgByteArr = io.BytesIO()           # pixel dimensions and font sizes are doubled to avoid pixelated
    font = ImageFont.truetype(font=os.path.join(onedrive_folder, 'LiberationSans-Regular.ttf'), size=16)
    bold_font = ImageFont.truetype(os.path.join(onedrive_folder, 'LiberationSans-Bold.ttf'), 20)
    text_image = Image.new(mode='RGB', size=size, color='#ffffff')
    draw = ImageDraw.Draw(im=text_image)
    # drawing headers
    headers = ['Product Name', 'Variation Name', 'Qty']
    draw.text(xy=(10, 24), text=headers[0], font=bold_font, fill='black', anchor='ls')
    draw.text(xy=(300, 24), text=headers[1], font=bold_font, fill='black', anchor='ls')
    draw.text(xy=(550, 24), text=headers[2], font=bold_font, fill='black', anchor='ls')
    draw.line((10, 30, 586, 30), fill='black', width=1)
    # drawing order details
    total_products, total_lines = 0, 3
    for index, order in order_details.iterrows():
        product_name = textwrap.wrap(order['Product Name'].split('//')[0], width=32)
        product_name_lines = '\n'.join(product_name[:3])
        variation_sku = [re.sub('.*:', '', option) for option in order['Product Variation'].split(',') if option != ''][:2]
        variation_sku.append(order['SKU'][:33])
        variation_sku_lines = '\n'.join(variation_sku)
        draw.text(xy=(10, 20*total_lines), text=product_name_lines, font=font, fill='black', anchor='ls')
        draw.text(xy=(300, 20*total_lines), text=variation_sku_lines, font=font, fill='black', anchor='ls')
        draw.text(xy=(560, 20*total_lines), text=str(int(order['Qty'])), font=font, fill='black', anchor='ls')
        total_lines += len(max(product_name[:3], variation_sku, key=len)) + 0.5
        total_products += 1
        if total_lines >= 40:
            break
    text_image.save(imgByteArr, format='PNG')
    return imgByteArr.getvalue(), total_products


def convert_packing_list():
    """Converts exported template excel from Ginee to customized Packing & Picking List in PDF"""
    print("Starting Ginee Packing List Converter")
    new_doc = fitz.open()

    # Finding latest downloaded excel
    list_of_excels = glob.glob(os.path.join(downloads_folder, '*.xlsx'))
    latest_file = max(list_of_excels, key=os.path.getctime)
    df = pd.read_excel(latest_file)
    df.fillna({'Product Variation': ''}, inplace=True)
    df = df[df['Product Status'] == 'Paid']                # filters out cancelled items
    order_nos = df.sort_values('SKU')['NO.'].unique()      # 1.0, 2.0, 3.0 ... 100.0

    picking_list = {}
    for order_no in order_nos:
        print(f"Processing Order No.: {order_no}")
        order = df[(df['NO.'] == order_no)]
        # creates page with a6 portrait
        a6_format = fitz.paper_rect('a6')
        new_page = new_doc.newPage(width = a6_format.width, height = a6_format.height)  # w, h = (298.0, 420.0)
        # inputs greeting & ty image
        greeting_name_img = draw_greetings(customer_name=order['Buyer Name'].values[0])
        new_page.insertImage((0, 20, 298, 70), stream=greeting_name_img, overlay=True)
        new_page.insertImage((0, 70, 298, 260), stream=ty_image_bytes)
        # insert order number & barcode
        order_number = order['Order ID'].values[0]
        barcode = generate_barcode(text=order_number, type='qrcode')
        new_page.insertImage((5, 3, 80, 20), stream=draw_text(order_number, (120, 30), 12))
        new_page.insertImage((255, 0, 293, 33), stream=barcode)
        # draws order details
        order_details_img = draw_order_details(order)[0]
        new_page.insertImage((0, 270, 298, 420), stream=order_details_img)
        # finally, inputs buyer note
        if all(order['Buyer Note'].notnull()):
            buyers_note = f"Buyer's Note: {order['Buyer Note'].values[0]}"
            new_page.insertImage((5, 401, 140, 418), stream=draw_text(buyers_note, (300, 30), 12))

    # Picking list
    print("Adding Picking List")
    table = pd.pivot_table(df, values=['Product Name', 'Product Variation', 'Qty', 'SKU'], index='Inventory SKU',
                                aggfunc={'Product Name':'first', 'Product Variation':'first', 'SKU':'first', 'Qty': np.sum})
    table.sort_values(by='SKU', inplace=True)
    print_date = dt.datetime.today().strftime('Print Date: %A, %b %d %Y')
    print_date_img = draw_text(print_date, (200, 30), 12)

    total_processed_products, page_no = 0, 0
    while total_processed_products != len(table):
        new_page = new_doc.newPage(pno = page_no, width = a6_format.width, height = a6_format.height)
        product_details = draw_order_details(table[total_processed_products: ], size=(596, 840))
        new_page.insertImage((0, 0, 298, 420), stream=product_details[0])
        total_processed_products += product_details[1]
        new_page.insertImage((200, 400, 290, 415), stream=print_date_img)
        page_no += 1

    print("Saving")
    save_location = os.path.join(onedrive_folder, 'Ginee Packing List.pdf')
    new_doc.save(save_location)
    new_doc.close()

    print("Opening PDF")
    webbrowser.open(save_location)


if __name__ == '__main__':
    convert_packing_list()