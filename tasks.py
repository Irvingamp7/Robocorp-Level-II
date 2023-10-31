from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive
import csv
import aspose.pdf as ap
import os
import time


@task
def insert_orders_to_system():
    """Insert the orders data to the system and export it as a PDF"""
    browser.configure(slowmo=1000,)
    open_the_intranet_website()
    click_ok()
    download_csv()
    read_the_csv_data()
    add_zip_file()
    pdf_cleanup()
    

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def click_ok():
    """confirmation click when entering the page"""
    page = browser.page()  
    page.click("text=OK")


def download_csv():
    """download csv file from given page"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def read_the_csv_data():
    """Read data from csv and fill in the sales form"""
    file = open('orders.csv')
    csvreader = csv.reader(file)
    encabezado = []
    encabezado = next(csvreader)
    print("**********************")
    print(encabezado[0])
    global rows
    rows = []
    for row in csvreader:
        rows.append(row)
    print("//////////////////////////")
    print(rows)

    file.close()
    #page = browser.page()
    for i in rows:
        fill_and_submit_robot(i)
        click_to_make_order()
        click_another_button()


def click_another_button():
    page = browser.page() 

    page.wait_for_selector('#order-another')
    page.click("#order-another")
    click_ok()
    

def fill_and_submit_robot(row):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()

    page.select_option("#head", str(row[1])) #head

    if row[2]=='1':
        page.check("#id-body-1",)
    elif row[2]=='2':
        page.check("#id-body-2",)
    elif row[2]=='3':
        page.check("#id-body-3",)
    elif row[2]=='4':
        page.check("#id-body-4",)
    elif row[2]=='5':
        page.check("#id-body-5",)
    else:
        page.check("#id-body-6",)

    #str(row[0][2])) #body
    page.fill('xpath=/html/body/div/div/div[1]/div/div[1]/form/div[3]/input', row[3])#legs
    page.fill("#address", str(row[4])) #addressP


def click_to_make_order():
    """click de creacion del robot parte final"""
    page = browser.page() 

    #if page.is_visible()
    #page.click("#preview")
    #page.wait_for_selector('#robot-preview-image')
    page.click("#order")
    if page.is_visible("#order-completion"):
        #page.wait_for_selector('#order-completion')
        capture_image_robot()
        export_as_pdf()
        
    elif page.is_visible("#root > div > div.container > div > div.col-sm-7 > div"):
            print("se fue un error")
            #page.click("#order")
            click_to_make_order()
    else:
        print("algo")

def capture_image_robot():
    """Take a screenshot of the page"""
    page = browser.page()
    global contenedor
    contenedor = page.text_content("#receipt > p.badge.badge-success")
    print(contenedor)
    page.screenshot(path=f"output/{contenedor}.png")

def export_as_pdf():
    """Export the data to pdf file"""
    page = browser.page()
    robot_info_html = page.locator("#order-completion").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(robot_info_html, f"output/{contenedor}.pdf")
    
    save_image_in_pdf_file(f"output/{contenedor}.pdf", f"output/{contenedor}1.pdf", f"output/{contenedor}.png")


def save_image_in_pdf_file(pdf_original, pdf_final, imagen_png):
    """save image to robot receipt"""
    # Open document
    document = ap.Document(pdf_original)
    document.pages[1].add_image(imagen_png, ap.Rectangle(100, 300,500, 700, True))
    document.save(pdf_final)

    
    # Verificar si el archivo existe antes de eliminarlo
    if os.path.exists(pdf_original):
        # Delete the file
        os.remove(pdf_original)
    else:
        print("it doesnt exists.")

    if os.path.exists(imagen_png):
        # Eliminar el archivo
        os.remove(imagen_png)
    else:
        print("it doesnt exists.")

def add_zip_file():
    """create the file in zip format to save it in the output"""
    time.sleep(8)
    share = os.getcwd()
    directory = share + '/output'
    contenido = os.listdir(directory)
    pdf_files = []
    for fichero in contenido:
        if os.path.isfile(os.path.join(directory, fichero)) and fichero.endswith('.pdf'):
            pdf_files.append(fichero)
    print(pdf_files)

    try:

        lib = Archive()
        lib.archive_folder_with_zip("./output", "output/orders.zip", exclude="*.robolog")
        time.sleep(5)
        lib.add_to_archive(pdf_files, "./output/orders.zip", "./output")

    except FileNotFoundError:
        print("irrelevante")    
    
    
def pdf_cleanup():
    time.sleep(5)
    files_dir = os.getcwd() + '/output'
    
    print(files_dir)
    delete_file = [f for f in os.listdir(files_dir) if f.endswith("1.pdf",)]
    print(delete_file)
    for f in delete_file:
        os.remove(f)
            