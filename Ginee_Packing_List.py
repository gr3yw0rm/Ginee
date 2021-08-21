import os
import re
import io
import time
import json
import fitz
import qrcode
import base64
import datetime as dt
import logging
from logging.handlers import RotatingFileHandler
import webbrowser
from win10toast_click import ToastNotifier 

# Folder locations
downloads_folder = os.path.join(os.environ.get('HOMEPATH'), 'Downloads')
onedrive_folder = os.path.join(os.getenv('HOMEPATH'), 'OneDrive', 'Shared Files - Shop', 'Python Scripts', 'Ginee')

# Logging Configuration
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
rotating_file_handler = RotatingFileHandler(filename=os.path.join(onedrive_folder, 'LOG.log'), mode='a', maxBytes=5000000)
log_formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s', datefmt='%d-%b-%y %H:%M:%S')
rotating_file_handler.setFormatter(log_formatter)
logger.addHandler(rotating_file_handler)


def add_barcode(doc):
    """Adds barcode in Ginee Packing List pdf file"""
    
    # doc = fitz.open(file_location)

    new_doc = fitz.open()

    logger.info(f"Total Orders: {doc.page_count}")

    for page in doc:
        if not page._isWrapped:
            page.wrap_contents()

        # Scales to a6 portrait
        fmt = fitz.paper_rect('a6')
        page = new_doc.newPage(width = fmt.width, height = fmt.height)
        print(f"Page number: {page.number}")
        page.showPDFpage(page.rect, doc, page.number)

        pdf_json = json.loads(page.get_text('json'))

        with open('Output.json', 'w') as f:
            f.write(page.get_text('json'))

        # PDF's Rect / Pixel location
        width, height = pdf_json['width'], pdf_json['height']

        for block in pdf_json['blocks']:
            try:
                # Order number is located every after image ~> [platform's image] Order Number
                if 'image' in block:
                    is_order_number = True
                    continue

                # Adds barcode according to order number's rect (w, h, w, h)
                elif is_order_number:
                    print(block)
                    text = block['lines'][0]['spans'][0]['text']
                    w, h, w1, h1 = block['lines'][0]['spans'][0]['bbox']
                    order_number = re.findall('(\w*\d*)', text)[0]
                    logger.info(f"Inputting barcode of Order # {order_number}")
                    qr = qrcode.QRCode()
                    qr.add_data(order_number)
                    img = qr.make_image(back_color='TransParent')
                    # img = qrcode.make(order_number)   # w/ white background
                    buffered = io.BytesIO()
                    img.save(buffered, format='PNG', transperancy=0, fill=(255, 0, 0))
                    rect = fitz.Rect(width/2-5, 3, width/2+15, 30)
                    page.insertImage(rect, stream=buffered.getvalue())

                is_order_number = False

            except NameError:
                continue

    new_doc.save(os.path.join(onedrive_folder, 'Ginee Packing List.pdf'))
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

    file_location = os.path.join(downloads_folder, 'Ginee_sample.pdf')
    os.chdir(downloads_folder)

    # initialize 
    toaster = ToastNotifier()

    while True:
        try:
            for filename in os.listdir(downloads_folder):

                if filename.endswith('.pdf'):

                    with fitz.open(filename) as doc:
                        first_page = doc.load_page(0)
                        matches = ['Buyers', 'Variation Name']

                        created_within_10secs = dt.datetime.now().timestamp() - os.path.getctime(filename) <= 10
                        print(f"{filename}: {all(match in first_page.get_text('json') for match in matches)} and {created_within_10secs}")
                        print(all(match in first_page.get_text('json') for match in matches) and created_within_10secs)
                        # Skips to next file if no matches found & created more than minute
                        if all(match in first_page.get_text('json') for match in matches) and created_within_10secs:
                            logger.info("Ginee Packing List Found!")
                            add_barcode(doc)
                            logger.info('Ginee Packing List Conversion Successful & Ready to Print')
                            # showcase
                            toaster.show_toast(
                                "Ginee Packing List", # title
                                "Click to print! >>", # message 
                                icon_path=os.path.join(onedrive_folder, 'ginee-app-logo.ico'), # 'icon_path' 
                                duration=15, # for how many seconds toast should be visible; None = leave notification in Notification Center
                                threaded=False, # True = run other code in parallel; False = code execution will wait till notification disappears 
                                callback_on_click=open_url(os.path.join(onedrive_folder, 'Ginee Packing List.pdf')) # click notification to run function 
                                )



                    # os.replace(filename, os.path.join('Ginee Packing Lists', filename))

            logger.warning("No Ginee Packing List Found")

        # Logs error
        except Exception as e:
            logger.error(e)

        time.sleep(3)


def test_toast_notification():
    toaster = ToastNotifier()
    toaster.show_toast(
    "Ginee Packing List", # title
    "Click to print! >>", # message 
    icon_path=os.path.join(onedrive_folder, 'ginee-app-logo.ico'), # 'icon_path' 
    duration=15, # for how many seconds toast should be visible; None = leave notification in Notification Center
    threaded=False, # True = run other code in parallel; False = code execution will wait till notification disappears 
    callback_on_click=open_url(os.path.join(onedrive_folder, 'Ginee Packing List.pdf')) # click notification to run function 
    )


if __name__ == '__main__':
    # test_toast_notification()
    main()