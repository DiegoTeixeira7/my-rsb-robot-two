# https://medium.com/@swapneil.basutkar/robocorp-level-2-python-rpa-bot-a8a6811545b2
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
    close_annoying_modal()
    download_csv_file()
    orders = get_orders()
    loop_orders(orders)
    archive_receipts()

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads a CSV file from the given URL."""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def close_annoying_modal():
    """Close annoying modal"""
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Read data from csv and fill in the sales form"""
    tables = Tables()
    worksheet = tables.read_table_from_csv(
        "orders.csv",
        header=True,
        columns=None,  # Você pode especificar os nomes das colunas aqui, se necessário
        dialect=None,  # Opcional: especificar o dialeto do CSV, como 'excel', 'excel-tab', ou 'unix'
        delimiters=None,  # Opcional: especificar os delimitadores possíveis como uma string
        column_unknown='Unknown',  # Nome da coluna para campos desconhecidos
        encoding=None  # Codificação de texto para o arquivo de entrada, usa a codificação do sistema por padrão
    )
    return worksheet

def fill_the_form(row):
    """Fills in the order data"""
    page = browser.page()

    page.select_option("#head", str(row["Head"]))
    page.set_checked("#id-body-"+str(row["Body"]), str(row["Body"]))
    page.fill("#address", row["Address"])
    page.fill("input[placeholder='Enter the part number for the legs']", row["Legs"])

    preview_robo()
    send_form(row)

def preview_robo():
    """preview robo"""
    page = browser.page()
    page.click("button:text('Preview')")

def loop_orders(orders):
    """Loop orders"""
    for row in orders:
        fill_the_form(row)

def send_form(row):
    """send form"""
    page = browser.page()
    while True:
        page.click("button:text('ORDER')")
        order_another = page.query_selector("#order-another")

        if order_another:
            pdf = store_receipt_as_pdf(row["Order number"])
            screenshot = screenshot_robot(row["Order number"])
            embed_screenshot_to_receipt(screenshot, pdf)
            page.click("button:text('ORDER ANOTHER ROBOT')")
            close_annoying_modal()
            break

def store_receipt_as_pdf(order_number):
    """store receipt as pdf"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf = PDF()
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """screenshot_robot"""
    page = browser.page()
    screenshot_path = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot to the bot receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot, 
                                   source_path=pdf_file, 
                                   output_path=pdf_file)

def archive_receipts():
    """Archives all the receipt pdfs into a single zip archive"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")