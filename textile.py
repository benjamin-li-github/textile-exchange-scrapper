from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import time
import csv


def initialize_driver():
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
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()), options=options
    )
    return driver


def open_website(driver, url):
    driver.get(url)
    time.sleep(5)


def load_dashboard(driver):
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
        print(f"Error loading dashboard: {e}")


def click_search_button(driver):
    try:
        search_elements__ = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "paddingDisabled"))
        )
        search_elements = WebDriverWait(search_elements__[-1], 5).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "ui-role-button-text"))
        )
        for search_element in search_elements:
            print(search_element.text)
            if search_element.text == "Search":
                search_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(search_elements__[-1])
                )
                search_btn.click()
                break
    except Exception as e:
        print(f"Error clicking search button: {e}")


def process_data(driver):
    data = []
    last_batch_data = []
    current_batch_data = []

    while len(data) < 55000:
        try:
            data__ = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "main-cell"))
            )
            i = 0 if len(data) == 0 else 8 + 8

            temp_data = []
            for i in range(i, len(data__)):
                if (i + 1) % 8 != 0:
                    temp_data.append(data__[i].text)
                else:
                    temp_data.append(data__[i].text)
                    data.append(temp_data)
                    current_batch_data.append(temp_data)
                    temp_data = []

            if last_batch_data != current_batch_data:
                last_batch_data = current_batch_data
                current_batch_data = []
            else:
                save_data_to_csv(data)
                return

            print(len(data))

        except Exception as e:
            print(f"Error processing data: {e}")
        driver.execute_script("arguments[0].scrollIntoView();", data__[-1])
        time.sleep(0.05)

    print("Complete")


def save_data_to_csv(data, filename="textile.csv"):
    with open(filename, "w", encoding="UTF-8", newline="") as f:
        write = csv.writer(f)
        write.writerows(data)


def main():
    driver = initialize_driver()
    try:
        open_website(driver, "https://textileexchange.org/find-certified-company/")
        load_dashboard(driver)
        click_search_button(driver)
        process_data(driver)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
