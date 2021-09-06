import os
import re
import io
import time
import json
import fitz
import base64
import qrcode
import barcode
from barcode.writer import ImageWriter
import datetime as dt
from PIL import Image
import logging
from logging.handlers import RotatingFileHandler
import webbrowser
from win10toast_click import ToastNotifier 


# Folder locations
downloads_folder = os.path.join(os.environ.get('HOMEPATH'), 'Downloads')
onedrive_folder = os.path.join(os.getenv('HOMEPATH'), 'OneDrive', 'Shared Files - Shop', 'Python Scripts', 'Ginee')

# Multiple package icon to bytes
with Image.open(os.path.join(onedrive_folder, 'package_icon.png'), 'r') as package_icon:
    imgByteArr = io.BytesIO()
    package_icon.save(imgByteArr, format='PNG')
    package_icon_bytes = imgByteArr.getvalue()

barcode.base.Barcode.default_writer_options['write_text'] = False

# Logging Configuration
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
stream_handler = logging.StreamHandler()
rotating_file_handler = RotatingFileHandler(filename=os.path.join(onedrive_folder, 'LOG.log'), mode='a', maxBytes=5000000)
log_formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s', datefmt='%d-%b-%y %H:%M:%S')
rotating_file_handler.setFormatter(log_formatter)
logger.addHandler(stream_handler)
logger.addHandler(rotating_file_handler)


def add_barcode(doc):
    """Adds barcode in Ginee Packing List pdf file"""
    
    # doc = fitz.open(file_location)

    new_doc = fitz.open()

    logger.info(f"Total Orders: {doc.page_count}")

    # tabulates data
    pdf_json = ""
    for page in doc:
        pdf_json += page.get_text('json')

    # Editing PDF
    for page in doc:
        if not page._isWrapped:
            page.wrap_contents()

        # Scales to a6 portrait
        fmt = fitz.paper_rect('a6')
        page = new_doc.newPage(width = fmt.width, height = fmt.height)
        print(f"Page number: {page.number}")
        page.showPDFpage(page.rect, doc, page.number)

        page_json = json.loads(page.get_text('json'))

        with open('Output.json', 'w') as f:
            f.write(page.get_text('json'))

        # PDF's Rect / Pixel location
        width, height = page_json['width'], page_json['height']

        for block in page_json['blocks']:
            try:
                # Order number is located every after image ~> [platform's image] Order Number
                if 'image' in block:
                    is_order_number = True
                    continue

                # Adds barcode according to order number's rect (w, h, w, h)
                elif is_order_number:
                    print(block)
                    text = block['lines'][0]['spans'][0]['text']
                    w0, h0, w1, h1 = block['lines'][0]['spans'][0]['bbox']
                    order_number = re.findall('(\w*\d*)', text)[0]
                    logger.info(f"Inputting barcode of Order # {order_number}")
                    
                    ## Adds barcode
                    # buffered = io.BytesIO()
                    # barcode.generate('Code39', order_number, writer=ImageWriter(), output=buffered)
                    # rect = fitz.Rect(w0+50, h0-3, w0+150, h1+3)
                    # page.insertImage(rect, stream=buffered.getvalue())

                    # Adds qrcode
                    qr = qrcode.QRCode()
                    qr.add_data(order_number)
                    img = qr.make_image(back_color='TransParent')
                    # img = qrcode.make(order_number)   # w/ white background
                    buffered = io.BytesIO()
                    img.save(buffered, format='PNG', transperancy=0, fill=(255, 0, 0))
                    rect = fitz.Rect(w0+60, h0-4, w0+90, h1+2)
                    page.insertImage(rect, stream=buffered.getvalue())

                    # Adds multiple package icon & skus
                    if len(re.findall(order_number, pdf_json)) > 1:
                        # icon
                        logger.debug(f"\tMultiple orders in one package found")
                        rect = fitz.Rect(w0+90, h0, w0+100, h1)
                        page.insertImage(rect, stream=package_icon_bytes)
                        # # sku
                        # sku = "sample-sku123456-uni(4)"

                        # rect_x1 = 260
                        # rect_y1 = h0-3
                        # rect_x2 = 297
                        # rect_y2 = h1 + 2

                        # rect_width = rect_x2 - rect_x1
                        # rect_height = rect_y2 - rect_y1

                        # rect = (rect_x1, rect_y1, rect_x2, rect_y2)

                        # fontsize_to_use = rect_width/len(sku)*2 + 0.2

                        # page.insertTextbox(rect, sku,
                        # fontsize=fontsize_to_use,
                        # fontname="Times-Roman",
                        # align=1)


                is_order_number = False

            except NameError:
                continue

    new_doc.save(os.path.join(onedrive_folder, 'Ginee Picking List.pdf'))
    return 


def open_url(page_url):
    try: 
        webbrowser.open_new(page_url)
        print('Opening URL...')  
    except: 
        print('Failed to open URL. Unsupported variable type.')


def main():
    """Script for converting downloaded Ginee Packing Lists"""
    logger.info("Starting Ginee Packing List Converter")

    file_location = os.path.join(onedrive_folder, 'Picking List.pdf')
    os.chdir(downloads_folder)

    # initialize 
    toaster = ToastNotifier()

    while True:
        try:
            for filename in os.listdir(downloads_folder):

                if filename.endswith('.pdf'):

                    with fitz.open(filename) as doc:
                        first_page = doc.load_page(0)
                        matches = ['Picking List', 'Print Date', 'Operator', 'Total Product']

                        created_within_10secs = dt.datetime.now().timestamp() - os.path.getctime(filename) <= 100000000
                        print(f"{filename}: {all(match in first_page.get_text('json') for match in matches)} and {created_within_10secs}")
                        print(all(match in first_page.get_text('json') for match in matches) and created_within_10secs)
                        # Skips to next file if no matches found & created more than minute
                        if all(match in first_page.get_text('json') for match in matches) and created_within_10secs:
                            logger.info("Ginee Picking List Found!")
                            add_barcode(doc)
                            logger.info('Ginee Picking List Conversion Successful & Ready to Print')
                            # showcase
                            print(file_location)
                            toaster.show_toast(
                                "Ginee Picking List", # title
                                "Click to print! >>", # message 
                                icon_path=os.path.join(onedrive_folder, 'ginee-app-logo.ico'), # 'icon_path' 
                                duration=15, # for how many seconds toast should be visible; None = leave notification in Notification Center
                                threaded=False, # True = run other code in parallel; False = code execution will wait till notification disappears 
                                callback_on_click=open_url(file_location) # click notification to run function 
                                )
                    # os.replace(filename, os.path.join('Ginee Picking Lists', filename))

            logger.warning("No Ginee Picking List Found")

        # Logs error
        except Exception as e:
            logger.error(e)

        time.sleep(3)


def pdf_to_json(file_location, file_name):
    with fitz.open(file_location) as doc:
        with open(file_name, 'w') as packing_list:
            # for page in doc:
            #     packing_list.write(page.get_text('json'))
            packing_list.write(doc.load_page(0).get_text('json'))

if __name__ == '__main__':
    # pdf_to_json(file_location=os.path.join(downloads_folder, 'Picking List.pdf'), file_name='Picking List.json')
    main()
    pass