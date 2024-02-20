import time
import pandas
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

excel_data_df = pandas.read_excel('data.xlsx', sheet_name='Sheet1')
number_of_row, number_of_column = excel_data_df.shape

BASE_URL = "https://messages.sideline.com/login"
MAX_WAIT_TIME = 20

# Elements xpath
phone_number_xpath = "//*[@id='number']"
password_xpath = "//*[@id='password']"
login_button_xpath = "//sc-button[normalize-space()='Login']"
login_error_message = "//div[contains(text(), 'Invalid phone number or password')]"
inbox_xpath = "//*[normalize-space()='Inbox']"
three_dot_menu = "//*[@id='main-content']/weblayout-page/ion-split-pane/ion-menu[1]/ion-content/conversation-list/ion-list/ion-item-sliding/ion-item/ion-label/div[2]/ion-buttons/ion-button/i"
sync_contact_dismiss = "//*[@id='SyncContactsXDismissPopup']"


# Helper function
def get_driver(browser_name: str) -> WebDriver:
    if browser_name.lower() == "chrome":
        return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    elif browser_name.lower() == "edge":
        return webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
    elif browser_name.lower() == "firefox":
        return webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
    else:
        return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))


def click_fill_text_by_js_script(browser_driver, element, user_input):
    location = element.location
    actions = ActionChains(driver)
    # Move to the specified coordinates and click
    actions.move_by_offset(location['x'], location['y']).click().send_keys(user_input).perform()
    return browser_driver


def write_to_txt_result(filename, content):
    with open(filename, "a") as file:
        if file.tell() > 0:
            file.write("\n")
        file.write(content)


def find_element_with_exception_xpath(
        driver: WebDriver,
        xpath_locator: str,
) -> WebElement:
    """
    Finds a web element using given XPath locator and returns it.

    Args:
        driver: The Selenium WebDriver instance (typing: WebDriver)
        xpath_locator: The XPath expression to locate the element (typing: str)

    Returns:
        The WebElement object if found, otherwise raises an appropriate exception.

    Raises:
        NoSuchElementException: If the element is not found.
        TimeoutException: If the WebDriverWait times out.
    """

    try:
        wait = WebDriverWait(driver, MAX_WAIT_TIME)  # Customize timeout as needed

        return wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath_locator))
        )

    except TimeoutException:
        raise TimeoutException(
            f"Element with XPath '{xpath_locator}' not found within the timeout!"
        )

    except NoSuchElementException:
        raise NoSuchElementException(
            f"Element with XPath '{xpath_locator}' not found!"
        )


def check_element_visibility(
        browser_driver: WebDriver,
        xpath_locator: str,
        timeout: int
) -> WebElement:
    try:
        wait = WebDriverWait(browser_driver, timeout=5)  # Customize timeout as needed

        return wait.until(
            EC.visibility_of_element_located((By.XPATH, xpath_locator))
        )

    except TimeoutException:
        print(f"Element with XPath '{xpath_locator}' not found within the timeout!")

    except NoSuchElementException:
        print(f"Element with XPath '{xpath_locator}' not found!")


def check_element_clickable(
        driver: WebDriver,
        xpath_locator: str,
) -> WebElement:
    try:
        wait = WebDriverWait(driver, MAX_WAIT_TIME)  # Customize timeout as needed

        return wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath_locator))
        )

    except TimeoutException:
        raise TimeoutException(
            f"Element with XPath '{xpath_locator}' not found within the timeout!"
        )

    except NoSuchElementException:
        raise NoSuchElementException(
            f"Element with XPath '{xpath_locator}' not found!"
        )


# print whole sheet data
print(number_of_row, number_of_column)
# print(excel_data_df.iloc[0, 1])


for row in range(1):
    try:
        phone_number = str(excel_data_df.iloc[row, 0])
        password = str(excel_data_df.iloc[row, 1])

        print(phone_number, password)

        driver = get_driver("chrome")  # browser name "chrome" or "firefox", "edge"
        # driver.maximize_window()
        driver.get(BASE_URL)
        actions = ActionChains(driver)

        input_phone_number: WebElement = find_element_with_exception_xpath(driver, phone_number_xpath)
        location_phone_number = input_phone_number.location
        actions.move_by_offset(location_phone_number['x'], location_phone_number['y']).click().send_keys(
            phone_number).perform()

        input_password: WebElement = find_element_with_exception_xpath(driver, password_xpath)
        location_password = input_password.location
        actions.move_to_element(input_password).click().send_keys(password).perform()

        login_button: WebElement = find_element_with_exception_xpath(driver, login_button_xpath)
        actions.move_to_element(login_button).click().perform()

        if check_element_visibility(driver, login_error_message, MAX_WAIT_TIME):
            driver.save_screenshot(f"login_screenshot//{row}_row_is_login_error.png")
            write_to_txt_result("login_result.csv", f"{phone_number}, {password}, Login Error")
            print(f"Row: {row} Number: {phone_number}, Password: {password}, Result: Login Error")
            driver.quit()
            continue
        if check_element_visibility(driver, inbox_xpath, MAX_WAIT_TIME):
            if check_element_visibility(driver, sync_contact_dismiss, 5):
                driver.find_element(By.XPATH, sync_contact_dismiss).click()
                time.sleep(2)

            three_dot_menu_button: WebElement = driver.find_element(By.XPATH, three_dot_menu)
            location_three_dot_menu = three_dot_menu_button.location

            actions.move_to_element(three_dot_menu_button).click().send_keys(password).perform()

            time.sleep(30)



            print(f"Row: {row} Number: {phone_number}, Password: {password}, Result: Passed")
            driver.save_screenshot(f"login_screenshot//{row}_row_is_passed.png")
            write_to_txt_result("login_result.csv", f"{phone_number}, {password}, Passed")
            driver.quit()
            continue
        else:
            print(f"Row: {row} Number: {phone_number}, Password: {password}, Result: Not able to determine")
            write_to_txt_result("login_result.csv", f"{phone_number}, {password}, Not able to determine")
            driver.save_screenshot(f"login_screenshot//{row}_row_is_not_determine.png")
            driver.quit()
            continue

    except Exception as e:
        # # Saving screenshot
        driver.save_screenshot(f"screenshot//login//exception//{row}_row_is_exception.png")
        # write_to_txt_result("login_result.csv", f"{phone_number}, {password}, Failed")
        # print(f"Row: {row} Number: {phone_number}, Password: {password}, Result: Failed")
        print(e.with_traceback())

        driver.quit()
