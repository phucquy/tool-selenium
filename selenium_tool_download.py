from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import time
import os


def tool_transaction_statisticsatic(username, password, hiden=False):
    dirname = os.path.dirname(__file__)
    path_download = os.path.join(dirname, 'XLS/')

    options = webdriver.ChromeOptions()
    if hiden == True:
        options.add_argument("--headless")  # Runs Chrome in headless mode.

    prefs = {"download.default_directory": path_download}
    options.add_experimental_option("prefs", prefs)
    options.add_argument('--no-sandbox')  # Bypass OS security model
    options.add_argument('start-maximized')  #
    options.add_argument('disable-infobars')
    options.add_argument("--disable-extensions")
    options.add_argument("--allow-running-insecure-content")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(15)
    driver.get("https://ib.techcombank.com.vn/servlet/BrowserServlet")

    # login Techcombank
    assert "Techcombank" in driver.title
    elem = driver.find_element_by_name("signOnName")
    elem.clear()
    elem.send_keys(username)
    elem = driver.find_element_by_name("password")
    elem.clear()
    elem.send_keys(password)
    driver.find_element_by_name("btn_login").click()
    time.sleep(3)

    # get statistical transaction techcombank
    try:
        account_transaction = driver.find_element_by_xpath("//div[@id='qw_top_menu']/ul/li[2]/a")
        account_transaction.click()
        dowload = driver.find_element_by_xpath("//table[@id='goButton']/tbody/tr/td/table/tbody/tr/td[2]")
        dowload.click()
        time.sleep(2)

        # Load new page
        current_window = driver.current_window_handle
        # get first child window
        new_window = driver.window_handles
        for w in new_window:
            if (w != current_window):
                driver.switch_to.window(w)
                time.sleep(5)
                break

        driver.execute_script("drilldown('1','1_1')")  # download file xls
        # driver.execute_script("drilldown('2','1_1')")  download file CSV
        time.sleep(5)
        print("Lấy dữ liệu thành công")
    except:
        print("Kết nối thất bại")
    finally:
        driver.quit()


if __name__ == '__main__':
    user_name = "0962265898"
    password = "Quy198"
    tool_transaction_statisticsatic(user_name, password)
