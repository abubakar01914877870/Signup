import time

import pandas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

excel_data_df = pandas.read_excel('data.xlsx', sheet_name='Sheet1')
number_of_row, number_of_column = excel_data_df.shape

BASE_URL = "https://my.sideline.com/#/login"
MAX_WAIT_TIME = 15
CC_NUMBER_1 = "4232"
CC_NUMBER_2 = "2320"
CC_NUMBER_3 = "4502"
CC_NUMBER_4 = "0279"
CC_EXP_MONTH = "02"
CC_EXP_YEAR = "26"
CC_CVV = "199"
BILLING_ZIP = "47404"

# Elements xpath
phone_number_xpath = "//input[@id='phoneNumber']"
next_button_xpath = "//button[@id='nextButton']"
trial_30_day_button_xpath = "//button[text()='Start a free 30-day trial']"
company_name_xpath = "//input[@id='company']"
first_name_xpath = "//input[@id='first']"
last_name_xpath = "//input[@id='last']"
email_xpath = "//input[@id='email']"
password_xpath = "//input[@id='password']"
freeTrial_button_xpath = "//button[@id='freeTrialButton']"
add_trial_ok_xpath = "//a[@id='freeTrialOK']"
cc_number_xpath = "//input[@id='card_number']"
cc_exp_xpath = "//input[@id='cc-exp']"
cc_cvv_xpath = "//input[@id='cc-csc']"
add_cc_xpath = "//span[text()='Add credit card']"
billing_zip_xpath = "//input[@id='billing-zip']"
message_cc_added_xpath = "//div[text()='Credit Card Added']"
errorTeamConversionMessage_xpath = "//div[@id='errorTeamConversionMessage']"


# Helper function
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
        driver: WebDriver,
        xpath_locator: str,
) -> WebElement:
    try:
        wait = WebDriverWait(driver, MAX_WAIT_TIME)  # Customize timeout as needed

        return wait.until(
            EC.visibility_of_element_located((By.XPATH, xpath_locator))
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


result = []
for row in range(2):
    try:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.maximize_window()
        driver.get(BASE_URL)
        phone_number = str(excel_data_df.iloc[row, 0])
        input_phone_number: WebElement = find_element_with_exception_xpath(driver, phone_number_xpath)
        input_phone_number.send_keys(phone_number)

        next_button: WebElement = find_element_with_exception_xpath(driver, next_button_xpath)
        next_button.click()

        trial_button: WebElement = find_element_with_exception_xpath(driver, trial_30_day_button_xpath)
        trial_button.click()

        company_name = str(excel_data_df.iloc[row, 2])
        company_name_input: WebElement = find_element_with_exception_xpath(driver, company_name_xpath)
        company_name_input.send_keys(company_name)

        first_name = str(excel_data_df.iloc[row, 3])
        first_name_input: WebElement = find_element_with_exception_xpath(driver, first_name_xpath)
        first_name_input.send_keys(first_name)

        last_name = str(excel_data_df.iloc[row, 4])
        last_name_input: WebElement = find_element_with_exception_xpath(driver, last_name_xpath)
        last_name_input.send_keys(last_name)

        email = str(excel_data_df.iloc[row, 5])
        email_input: WebElement = find_element_with_exception_xpath(driver, email_xpath)
        email_input.send_keys(email)

        password = str(excel_data_df.iloc[row, 1])
        password_input: WebElement = find_element_with_exception_xpath(driver, password_xpath)
        password_input.send_keys(password)

        freeTrial_button: WebElement = find_element_with_exception_xpath(driver, freeTrial_button_xpath)
        freeTrial_button.click()

        freeTrial_ok: WebElement = find_element_with_exception_xpath(driver, add_trial_ok_xpath)
        freeTrial_ok.click()

        iframe = WebDriverWait(driver, MAX_WAIT_TIME).until(
            EC.visibility_of_element_located((By.NAME, "stripe_checkout_app"))
        )
        driver.switch_to.frame(iframe)

        cc_number_input: WebElement = find_element_with_exception_xpath(driver, cc_number_xpath)
        cc_number_input.send_keys(CC_NUMBER_1)
        cc_number_input.send_keys(CC_NUMBER_2)
        cc_number_input.send_keys(CC_NUMBER_3)
        cc_number_input.send_keys(CC_NUMBER_4)

        cc_exp_input: WebElement = find_element_with_exception_xpath(driver, cc_exp_xpath)
        cc_exp_input.click()
        cc_exp_input.send_keys(CC_EXP_MONTH)
        cc_exp_input.send_keys(CC_EXP_YEAR)

        cc_cvv_input: WebElement = find_element_with_exception_xpath(driver, cc_cvv_xpath)
        cc_cvv_input.send_keys(CC_CVV)

        billing_zip_input: WebElement = find_element_with_exception_xpath(driver, billing_zip_xpath)
        billing_zip_input.send_keys(BILLING_ZIP)

        add_cc_button: WebElement = find_element_with_exception_xpath(driver, add_cc_xpath)
        add_cc_button.click()

        driver.switch_to.parent_frame()

        message_cc_added = check_element_visibility(driver, message_cc_added_xpath)

        freeTrial_ok_2: WebElement = check_element_visibility(driver, add_trial_ok_xpath)
        freeTrial_ok_2.click()

        errorTeamConversionMessage = WebDriverWait(driver, MAX_WAIT_TIME).until(
            EC.visibility_of_element_located((By.XPATH, errorTeamConversionMessage_xpath))
        )

        result.append({"Result": "Passed", "Number": phone_number})
        time.sleep(3)
        driver.quit()
    except Exception as e:
        # Saving screenshot
        driver.save_screenshot(f"screenshot//{row}_row_is_failed.png")
        result.append({"Result": "Failed", "Number": phone_number})

        print(f"Row number {row} is failed")
        print(e)

        driver.quit()
# Updating excel
output = pd.DataFrame(result, columns=["Number", "Result"])
output.to_excel("output.xlsx")
