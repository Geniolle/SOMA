#!/usr/bin/env python3
"""
Inspecciona o modal/popup aberto no Chrome e colecta os XPaths dos elementos críticos.
"""
import json
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')

def get_xpath(driver, element):
    """Retorna o XPath de um elemento."""
    try:
        xpath = driver.execute_script("""
            function getElementXPath(element) {
                if (element.id !== '')
                    return "//*[@id='" + element.id + "']";
                if (element === document.body)
                    return element.tagName.toLowerCase();
                var ix = 0;
                var siblings = element.parentNode.childNodes;
                for (var i = 0; i < siblings.length; i++) {
                    var sibling = siblings[i];
                    if (sibling === element)
                        return getElementXPath(element.parentNode) + "/" + element.tagName.toLowerCase() + "[" + (ix + 1) + "]";
                    if (sibling.nodeType === 1 && sibling.tagName.toLowerCase() === element.tagName.toLowerCase())
                        ix++;
                }
            }
            return getElementXPath(arguments[0]);
        """, element)
        return xpath
    except Exception as e:
        logging.error(f"Erro ao calcular XPath: {e}")
        return None

def find_element_info(driver, selector_type, selector_value, name):
    """Encontra um elemento e retorna informações sobre ele."""
    try:
        element = driver.find_element(selector_type, selector_value)
        xpath = get_xpath(driver, element)
        tag = element.tag_name
        text = element.text[:100] if element.text else ""
        visible = element.is_displayed()
        enabled = element.is_enabled()

        info = {
            "name": name,
            "xpath": xpath,
            "tag": tag,
            "text": text,
            "visible": visible,
            "enabled": enabled,
            "selector_used": f"{selector_type}={selector_value}"
        }
        logging.info(f"✓ {name} encontrado: {xpath}")
        return info
    except NoSuchElementException:
        logging.warning(f"✗ {name} NÃO encontrado com {selector_type}={selector_value}")
        return None
    except Exception as e:
        logging.error(f"✗ Erro ao procurar {name}: {e}")
        return None

def main():
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('user-data-dir=C:\\Users\\clayton.silva\\AppData\\Local\\Google\\Chrome\\User Data')
    options.add_argument('profile-directory=Default')

    try:
        # Tenta conectar ao Chrome já aberto
        logging.info("Conectando ao Chrome aberto...")
        driver = webdriver.Chrome(options=options)

        logging.info(f"URL atual: {driver.current_url}")

        # Aguarda um pouco para a página carregar
        import time
        time.sleep(2)

        # Lista de elementos para procurar
        search_patterns = [
            # Modais/Popups
            ("xpath", "//div[@role='dialog']", "Dialog (role=dialog)"),
            ("xpath", "//div[@role='alertdialog']", "Alert Dialog (role=alertdialog)"),
            ("xpath", "//div[contains(@class, 'modal')]", "Modal (class=modal)"),
            ("xpath", "//div[contains(@class, 'popup')]", "Popup (class=popup)"),
            ("xpath", "//div[contains(@class, 'swal')]", "SweetAlert (class=swal)"),

            # Botões OK/Confirmar
            ("xpath", "//button[contains(., 'OK')]", "Botão OK (texto)"),
            ("xpath", "//button[contains(., 'Sim')]", "Botão Sim (texto)"),
            ("xpath", "//button[contains(., 'Confirmar')]", "Botão Confirmar (texto)"),
            ("xpath", "//button[@class='swal2-confirm']", "SweetAlert Confirm Button"),
            ("xpath", "//span[contains(., 'OK')]//ancestor::button", "Botão OK (span)"),

            # Campos do modal de pagamento
            ("xpath", "//input[@name='num_documento']", "Nº Documento (name=num_documento)"),
            ("xpath", "//input[@name='numero_documento']", "Nº Documento (name=numero_documento)"),
            ("xpath", "//input[contains(@placeholder, 'Documento')]", "Nº Documento (placeholder)"),
            ("xpath", "//label[contains(., 'Nº Documento')]/following::input[1]", "Nº Documento (label)"),

            # Botão Salvar Pagamento
            ("xpath", "//button[contains(., 'Salvar')]", "Botão Salvar (texto)"),
            ("xpath", "//button[@id='botao_pagamento']", "Botão Salvar (id=botao_pagamento)"),
            ("xpath", "//button[contains(@class, 'salvar')]", "Botão Salvar (class)"),

            # Data Pagamento
            ("xpath", "//input[@name='data_pagamento']", "Data Pagamento (name)"),
            ("xpath", "//input[contains(@placeholder, 'Data')]", "Data Pagamento (placeholder)"),
        ]

        results = []
        logging.info("=" * 60)
        logging.info("Procurando elementos...")
        logging.info("=" * 60)

        for selector_type, selector_value, name in search_patterns:
            info = find_element_info(driver, selector_type, selector_value, name)
            if info:
                results.append(info)

        # Salvar resultados
        output_file = "C:\\workspace\\SOMA\\elementos_inspeccionados.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        logging.info("=" * 60)
        logging.info(f"✓ Resultados salvos em: {output_file}")
        logging.info(f"✓ {len(results)} elementos encontrados")
        logging.info("=" * 60)

        # Imprimir um resumo
        print("\n" + "=" * 60)
        print("RESUMO DOS ELEMENTOS ENCONTRADOS")
        print("=" * 60)
        for r in results:
            print(f"\n{r['name']}")
            print(f"  XPath: {r['xpath']}")
            print(f"  Tag: {r['tag']}")
            if r['text']:
                print(f"  Text: {r['text']}")
            print(f"  Visível: {r['visible']}, Habilitado: {r['enabled']}")

        # Tirar screenshot
        screenshot_file = "C:\\workspace\\SOMA\\artifacts\\screenshots\\inspeccionado_elementos.png"
        driver.save_screenshot(screenshot_file)
        logging.info(f"✓ Screenshot salvo: {screenshot_file}")

    except Exception as e:
        logging.error(f"Erro fatal: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Não fechar o driver, deixar o Chrome aberto
        logging.info("Script concluído. Deixando Chrome aberto para mais inspecção.")

if __name__ == '__main__':
    main()
