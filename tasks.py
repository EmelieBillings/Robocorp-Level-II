from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )

    open_robot_order_website()
    download_CSV_file()
    worksheet = get_orders()
    close_annoying_modal()
    fill_the_form(worksheet)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_CSV_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    csv_file = Tables()
    worksheet = csv_file.read_table_from_csv("orders.csv", header=True)
    return worksheet

def fill_the_form(worksheet):
    page = browser.page()
    for order in worksheet:
        page.select_option("#head", order['Head'])
        page.click("text="+body(order['Body']))
        page.fill("input[placeholder='Enter the part number for the legs']", order['Legs'])
        page.fill("input[placeholder='Shipping address']", order['Address'])
        page.click("button:text('Preview')")
        page.click("button:text('ORDER')")
        while page.is_visible("text=Error") or page.is_visible("text=Server") or page.is_visible("text=tried") :
            page.click("button:text('ORDER')")
        embed_screenshot_to_receipt(screenshot_robot(order['Order number']), store_receipt_as_pdf(order['Order number']))
        page.click("button:text('Order another robot')")
        close_annoying_modal()



def body(id):
    if(id == '1'):
        return "Roll-a-thor body"
    elif(id == '2'):
        return "Peanut crusher body"
    elif(id == '3'):
        return "D.A.V.E body"
    elif(id == '4'):
        return "Andy Roid body"
    elif(id == "5"):
        return "Spanner mate body"
    else:
        return "Drillbit 2000 body"

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Yep')")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(receipt, "output/receipts/sales_receipt_" + order_number + ".pdf")
    return "output/receipts/sales_receipt_" + order_number + ".pdf"

def screenshot_robot(order_number):
    page = browser.page()
    receipt = page.locator("#robot-preview-image")

    receipt.screenshot(path="output/receipt_screenshot_" + order_number + ".png")

    return "output/receipt_screenshot_" + order_number + ".png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.open_pdf(pdf_file)
    pdf.add_watermark_image_to_pdf(screenshot, pdf_file)
    pdf.close_pdf

def archive_receipts():
    zip = Archive()
    zip.archive_folder_with_zip('./output/receipts', 'output/receipts.zip')

    

