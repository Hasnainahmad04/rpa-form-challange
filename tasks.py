from robocorp.tasks import task

from RPA.Browser.Selenium import Selenium, By
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

WEBSITE_URL = "https://rpachallenge.com/"
EXCEL_FILE_URL = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"

selenium = Selenium()
http = HTTP()
files = Files()

@task
def rpa_form_challenge():
    """Main task function to complete the RPA form challenge.
    
    - Opens the challenge webpage.
    - Downloads the Excel file containing persons' information.
    - Fills out and submits the form for each person in the Excel file.
    - Closes the browser after completion.
    """
    open_challenge_page()
    download_persons_data_and_start_challenge()
    persons = get_persons_information()

    for person in persons:
        fill_the_form(person)
    
    selenium.close_all_browsers()

def open_challenge_page():
    """Opens the RPA challenge webpage in a new browser session."""
    selenium.open_available_browser()
    selenium.go_to(WEBSITE_URL)

def download_persons_data_and_start_challenge():
    """Downloads the Excel file containing persons' data and starts the challenge."""
    http.download(EXCEL_FILE_URL, 'output/person_information.xlsx', overwrite=True)
    selenium.click_button("Start")

def fill_the_form(data):
    """Fills and submits the form on the webpage with data from a dictionary.
    
    Args:
        data (dict): A dictionary containing form field names as keys 
                     and the corresponding data to fill as values.
                     
    The form fields are mapped by their 'ng-reflect-name' attributes.
    """
    form_mapping = {
        "labelPhone": "Phone Number",
        "labelFirstName": "First Name",
        "labelLastName": "Last Name",
        "labelEmail": "Email",
        "labelRole": "Role in Company",
        "labelAddress": "Address",
        "labelCompanyName": "Company Name"
    }
    for key, value in form_mapping.items():
        selenium.input_text(f"//*[@ng-reflect-name='{key}']", data[value])
    
    selenium.click_button('Submit')

def get_persons_information():
    """Reads and extracts persons' information from the Excel file.
    
    Returns:
        list: A list of dictionaries, where each dictionary contains data for a single person.
    """
    files.open_workbook('output/person_information.xlsx')
    worksheet = files.read_worksheet_as_table(header=True)
    files.close_workbook()
    return worksheet
