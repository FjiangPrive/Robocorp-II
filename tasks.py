from robocorp.tasks import task  # Importeer de taakdecorateur van de robocorp-takenmodule
from robocorp import browser  # Importeer de browsermodule van robocorp
from RPA.HTTP import HTTP  # Importeer de HTTP-module van RPA
from RPA.Tables import Tables  # Importeer de Tabellen-module van RPA
from RPA.PDF import PDF  # Importeer de PDF-module van RPA
from RPA.Archive import Archive  # Importeer de Archief-module van RPA
import shutil  # Importeer de shutil-module voor bestandsbewerkingen

@task
def order_robots_from_RobotSpareBin():
    """
    Voert de taak uit om robots te bestellen bij RobotSpareBin Industries Inc.
    Slaat het HTML-bestelontvangstbewijs op als een PDF-bestand.
    Maakt een screenshot van de bestelde robot.
    Voegt het screenshot van de robot toe aan het PDF-ontvangstbewijs.
    Maakt een ZIP-archief van de ontvangstbewijzen en de afbeeldingen.
    """
    browser.configure(
        slowmo=200,)  # Configureer de browser om langzamer te werken voor betere zichtbaarheid
    open_robot_order_website()  # Open de website voor het bestellen van robots
    download_orders_file()  # Download het bestand met bestellingen
    fill_form_with_csv_date() # Leest gegevens uit een CSV-bestand en vult het robotbestelformulier in met de gelezen gegevens.
    archive_receipts()  # Archiveer de ontvangstbewijzen
    clean_up()  # Maak de tijdelijke bestanden op

def open_robot_order_website():
    """Ga naar de website voor het bestellen van robots en accepteer pop-ups."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click('text=OK')  # Klik op de knop 'OK' om eventuele pop-ups te accepteren

def download_orders_file():
    """Download het bestand met bestellingen van de opgegeven URL."""
    http = HTTP()
    http.download(" https://robotsparebinindustries.com/orders.csv", overwrite=True)

def order_another_bot():
    """
    Klik op de knop 'Order Another' om een nieuwe robot te bestellen.

    Args:
        None

    Returns:
        None
    """
    page = browser.page()  # Haal de huidige pagina op van de browser
    page.click("#order-another")  # Klik op de knop 'Order Another' om een nieuwe robot te bestellen

def clicks_ok():
    """
    Klik op de knop 'OK' om een nieuwe bestelling voor robots te bevestigen.

    Args:
        None

    Returns:
        None
    """
    page = browser.page()  # Haal de huidige pagina op van de browser
    page.click('text=OK')  # Klik op de knop 'OK' om een nieuwe bestelling voor robots te bevestigen

def fill_and_submit_robot_data(order):
    """
    Vul de details van de robotbestelling in en klik op de 'Order' knop.

    Args:
        order (dict): Een dictionary met de details van de robotbestelling.
                      Moet de volgende sleutels bevatten:
                      - "Head": Een string die overeenkomt met het gewenste type hoofd.
                      - "Body": Een integer die overeenkomt met het gewenste type lichaam.
                      - "Legs": Een string met het onderdeelnummer voor de benen.
                      - "Address": Een string met het afleveradres van de bestelling.
                      - "Order number": Een integer met het bestelnummer.

    Returns:
        None
    """
    page = browser.page()  # Haal de huidige pagina op van de browser

    # Mapping van hoofdnummers naar hoofdnamen
    head_names = {
        "1": "Roll-a-thor head",
        "2": "Peanut crusher head",
        "3": "D.A.V.E head",
        "4": "Andy Roid head",
        "5": "Spanner mate head",
        "6": "Drillbit 2000 head"
    }

    head_number = order["Head"]  # Haal het nummer van het gewenste hoofd op uit de bestelling
    page.select_option("#head", head_names.get(head_number))  # Selecteer het juiste hoofdtype op de pagina

    page.click('//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{0}]/label'.format(order["Body"]))  # Selecteer het gewenste lichaamstype op de pagina
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])  # Vul het onderdeelnummer voor de benen in op de pagina
    page.fill("#address", order["Address"])  # Vul het afleveradres in op de pagina

    # Blijf op de 'Order' knop klikken totdat er geen extra bestellingen meer zijn
    while True:
        page.click("#order")
        order_another = page.query_selector("#order-another")
        if order_another:
            pdf_path = store_receipt_as_pdf(int(order["Order number"]))  # Sla de ontvangstbevestiging op als PDF
            screenshot_path = screenshot_robot(int(order["Order number"]))  # Maak een screenshot van de robotconfiguratie en sla het op
            embed_screenshot_to_receipt(screenshot_path, pdf_path)  # Voeg het screenshot toe aan de ontvangstbevestiging (PDF)
            order_another_bot()  # Bestel een andere robot met behulp van een bot
            clicks_ok()  # Bevestig het dialoogvenster voor het bestellen van een andere robot
            break  # Stop de loop

def store_receipt_as_pdf(order_number):
    """
    Slaat de ontvangstbevestiging van de robotbestelling op als een PDF-bestand.

    Args:
        order_number (int): Het bestelnummer van de robot.

    Returns:
        str: Het pad naar het opgeslagen PDF-bestand.
    """
    page = browser.page()  # Haal de huidige pagina op van de browser
    order_receipt_html = page.locator("#receipt").inner_html()  # Haal de HTML-code van de ontvangstbevestiging op van de pagina
    pdf = PDF()  # Maak een PDF-object aan
    pdf_path = "output/receipts/{0}.pdf".format(order_number)  # Bepaal het pad waar het PDF-bestand moet worden opgeslagen
    pdf.html_to_pdf(order_receipt_html, pdf_path)  # Converteer de HTML-code van de ontvangstbevestiging naar een PDF-bestand
    return pdf_path  # Geef het pad naar het opgeslagen PDF-bestand terug

def screenshot_robot(order_number):
    """
    Maakt een screenshot van de afbeelding van de bestelde robot.

    Args:
        order_number (int): Het bestelnummer van de robot.

    Returns:
        str: Het pad naar het opgeslagen screenshotbestand.
    """
    page = browser.page()  # Haal de huidige pagina op van de browser
    screenshot_path = "output/screenshots/{0}.png".format(order_number)  # Bepaal het pad waar het screenshotbestand moet worden opgeslagen
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)  # Maak een screenshot van de afbeelding van de bestelde robot
    return screenshot_path  # Geef het pad naar het opgeslagen screenshotbestand terug

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    """
    Voegt het screenshot toe aan de ontvangstbevestiging van de robot.

    Args:
        screenshot_path (str): Het pad naar het opgeslagen screenshotbestand.
        pdf_path (str): Het pad naar het PDF-bestand van de ontvangstbevestiging.

    Returns:
        None
    """
    pdf = PDF()  # Maak een PDF-object aan
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path,
                                    source_path=pdf_path,
                                    output_path=pdf_path)  # Voeg het screenshot toe aan de ontvangstbevestiging van de robot

def fill_form_with_csv_date():
    """Lees gegevens uit een CSV-bestand en vul het robotbestelformulier in."""
    csv_file = Tables()  # Maak een nieuw object aan om CSV-bestanden te lezen
    robot_orders = csv_file.read_table_from_csv("orders.csv")  # Lees de gegevens van robotbestellingen uit het CSV-bestand

    for order in robot_orders:
        fill_and_submit_robot_data(order)  # Vul het bestelformulier in en verstuur de bestelling

def archive_receipts():
    """
    Archiveert alle ontvangstbewijzen (PDF's) in één ZIP-archief.

    Args:
        None

    Returns:
        None
    """
    lib = Archive()  # Maak een archiefobject aan
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")  # Archiveer de map met ontvangstbewijzen in een ZIP-archief

def clean_up():
    """
    Maakt de tijdelijke mappen waar ontvangstbewijzen en screenshots worden opgeslagen schoon.

    Args:
        None

    Returns:
        None
    """
    shutil.rmtree("./output/receipts")  # Verwijder de map met ontvangstbewijzen
    shutil.rmtree("./output/screenshots")  # Verwijder de map met screenshots
