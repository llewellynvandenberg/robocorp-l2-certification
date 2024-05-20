from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables

from time import sleep
import zipfile
import os


@task
def order_robots_from_RobotSpareBin():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    get_orders()
    click_accept()
    order_robots()
    
    
def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
def click_accept():
    """Presses the Yep button"""
    page = browser.page()
    page.click("text=Yep")
    
def fill_robot_info(row):
    page = browser.page()
    
    page.select_option("#head", row["Head"])
    page.click(f"#id-body-{row['Body']}")
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[3]/input", row['Legs'])
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[4]/input", row["Address"])
    
def get_preview(order_number):
    path = f'screenshots/{order_number}.png'
    page = browser.page()
    page.click('#preview')
    sleep(1)
    page.locator('#robot-preview-image').screenshot(path= path)
    return path
    
def get_receipt(preview_path, order_number):
    page = browser.page()
    path = f'pdfs/{order_number}.pdf'
    
    while 1:
        try:
            page.click('#order')
            receipt_html = page.locator("#receipt").inner_html()
            break
        except:
            print('error')
            pass
            
    

    pdf = PDF()
    pdf.html_to_pdf(receipt_html, path)
    pdf.add_watermark_image_to_pdf(image_path = preview_path, source_path = path, output_path=path)
    
    #pdf.add_files_to_pdf(files = [preview_path], target_document= path, append = True)
    
    page.click('#order-another')
    click_accept()



def order_robots():
    """Read data from excel and fill in the sales form"""

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )

    for row in orders:
        fill_robot_info(row)
        tmp = get_preview(row['Order number'])
        get_receipt(tmp, row['Order number'])
        
    zip_files('pdfs', 'output/orders.zip')
        

        
def zip_files(directory, output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(directory):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)