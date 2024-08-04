from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import time
import csv


def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--incognito")
    options.add_argument("--enable-javascript")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-extensions")
    options.add_argument("--dns-prefetch-disable")
    options.add_argument("--disable-infobars")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    # options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()), options=options
    )
    return driver


def read_sc_numbers(filename):
    sc_numbers = []
    with open(filename, "r", encoding="UTF-8", newline="") as f:
        csvdata = csv.reader(f)
        for row in csvdata:
            sc_numbers.append(row[5])
    return sc_numbers


def initialize_data():
    return {
        "contact_data": [
            [
                "SC Number",
                "Contact",
                "Address",
                "State/Province",
                "Country/Area",
                "Public Email",
                "Website",
            ]
        ],
        "scope_certification_data": [
            "SC Number",
            "Number",
            "Version Number",
            "Organization Name",
            "Issue Date",
            "Valid Until",
            "Program",
            "Standard",
        ],
        "facility_data": [
            [
                "SC Number",
                "Type",
                "Name",
                "Address",
                "State/Province",
                "Country/Area",
                "Standard",
                "Process Category Code",
                "Process Category Description",
            ]
        ],
        "product_data": [
            [
                "SC Number",
                "Facility Name",
                "Product Category",
                "Category Code",
                "Detail Code",
                "Certified Material Code",
                "Material Description",
                "Raw Material %",
            ]
        ],
    }


def navigate_to_iframe(driver):
    try:
        dashboard_elm = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "block-emded-powerbi"))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView(true); window.scrollBy(0, -35);", dashboard_elm
        )
        time.sleep(5)
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
    except Exception as e:
        print(f"Error navigating to iframe: {e}")


def perform_search(driver, sc_number):
    try:
        search_area = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "text-filter-search"))
        )
        search_field = search_area.find_element(By.TAG_NAME, "input")
        search_field.click()
        search_field.send_keys(sc_number)
        time.sleep(1)
        search_field.click()
        search_field.send_keys(Keys.ENTER)
    except Exception as e:
        print(f"Error performing search: {e}")


def click_search_button(driver):
    try:
        driver.switch_to.default_content()
        dashboard_elm = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "block-emded-powerbi"))
        )
        driver.execute_script("arguments[0].scrollIntoView(false);", dashboard_elm)
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        search_elements__ = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "paddingDisabled"))
        )
        search_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((search_elements__[-1]))
        )
        search_btn.click()
    except Exception as e:
        print(f"Error clicking search button: {e}")


def select_sc_number(driver, sc_number):
    try:
        driver.switch_to.default_content()
        dashboard_elm = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "block-emded-powerbi"))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView(true); window.scrollBy(0, -35);", dashboard_elm
        )
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        data__ = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "main-cell"))
        )
        for data in data__:
            if data.text == sc_number:
                data.click()
                break
    except Exception as e:
        print(f"Error selecting SC number: {e}")


def click_get_additional_details(driver):
    try:
        driver.switch_to.default_content()
        dashboard_elm = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "block-emded-powerbi"))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView(true); window.scrollBy(0, -35);", dashboard_elm
        )
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        search_elements__ = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "paddingDisabled"))
        )
        search_elements__[1].click()
    except Exception as e:
        print(f"Error clicking 'Get Additional Details': {e}")


def extract_data(driver, sc_number, data_dict):
    try:
        datablocks = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "innerContainer"))
        )
        for idx, datablock in enumerate(datablocks):
            mid_viewport = WebDriverWait(datablock, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "mid-viewport"))
            )
            mid_viewport_rows = WebDriverWait(mid_viewport, 5).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "row"))
            )
            for mid_viewport_row in mid_viewport_rows:
                temp_data = [sc_number]
                mid_viewport_row_datacells = WebDriverWait(mid_viewport_row, 5).until(
                    EC.presence_of_all_elements_located(
                        (By.CLASS_NAME, "pivotTableCellWrap")
                    )
                )
                for mid_viewport_row_datacell in mid_viewport_row_datacells:
                    temp_data.append(mid_viewport_row_datacell.text)
                if idx == 0:
                    data_dict["contact_data"].append(temp_data)
                elif idx == 1:
                    data_dict["scope_certification_data"].append(temp_data)
                elif idx == 2:
                    data_dict["facility_data"].append(temp_data)
                elif idx == 3:
                    data_dict["product_data"].append(temp_data)
    except Exception as e:
        print(f"Error extracting data: {e}")


def click_back_button(driver):
    try:
        driver.switch_to.default_content()
        dashboard_elm = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "block-emded-powerbi"))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView(true); window.scrollBy(0, -60);", dashboard_elm
        )
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
        )
        search_elements__ = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "paddingDisabled"))
        )
        search_elements__[0].click()
    except Exception as e:
        print(f"Error clicking back button: {e}")


def save_to_csv(data_dict):
    with open("textile_contact.csv", "w", encoding="UTF-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(data_dict["contact_data"])

    with open("textile_scope_cert.csv", "w", encoding="UTF-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(data_dict["scope_certification_data"])

    with open("textile_facility.csv", "w", encoding="UTF-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(data_dict["facility_data"])

    with open("textile_product.csv", "w", encoding="UTF-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(data_dict["product_data"])


def main():
    driver = setup_driver()
    sc_numbers = read_sc_numbers("textile_1stlayer_complete.csv")
    data_dict = initialize_data()

    driver.get("https://textileexchange.org/find-certified-company/")
    time.sleep(3)

    navigate_to_iframe(driver)

    for i, sc_number in enumerate(sc_numbers):
        restart = False
        while not restart and i < len(sc_numbers):
            try:
                perform_search(driver, sc_number)
                click_search_button(driver)
                select_sc_number(driver, sc_number)
                click_get_additional_details(driver)
                extract_data(driver, sc_number, data_dict)
                click_back_button(driver)
                time.sleep(0.25)
            except Exception as e:
                print(f"Error processing SC number {sc_number}: {e}")
                restart = True
                driver.quit()
                driver = setup_driver()
                driver.get("https://textileexchange.org/find-certified-company/")
                navigate_to_iframe(driver)

    save_to_csv(data_dict)
    driver.quit()


if __name__ == "__main__":
    main()
