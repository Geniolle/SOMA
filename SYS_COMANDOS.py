import time
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        elemento = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
        return elemento
    except Exception as e:
        return None


def enviar_para_campo(valor_a_enviar: str, xpath: str, driver: webdriver) -> bool:
    descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, xpath)
    if descricao_campo:
        descricao_campo.send_keys(valor_a_enviar)
        return True
    else:
        time.sleep(2)
        return False

