###################################################################################
# selenium_scripts/navegacao.py
###################################################################################

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time

def tentar_encontrar_elemento(driver, by, value, timeout=5):
    for _ in range(timeout):
        try:
            return driver.find_element(by, value)
        except NoSuchElementException:
            time.sleep(1)
    return None

def clicar_no_elemento(driver, by, value, timeout=5):
    elemento = tentar_encontrar_elemento(driver, by, value, timeout)
    if elemento:
        elemento.click()
        return True
    return False

def redirecionar_para_inicio(driver):
    driver.get("https://verbodavida.info/apps/index.php")
    time.sleep(2)
