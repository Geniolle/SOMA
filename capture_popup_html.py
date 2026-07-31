#!/usr/bin/env python3
"""
Captura o HTML do popup que aparece após clicar 'Salvar Pagamento'
e salva num arquivo JSON para análise.
"""
import json
import logging
import time
import subprocess
import os
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')

def main():
    # Iniciar SOMA em background
    env = os.environ.copy()
    env['HEADLESS'] = 'false'

    logging.info("Iniciando SOMA em background...")
    soma_proc = subprocess.Popen(
        ['python', 'main.py'],
        cwd='C:\\workspace\\SOMA',
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )

    # Aguardar chegar ao ponto crítico (40-50 segundos)
    logging.info("Aguardando SOMA chegar ao ponto do erro (50s)...")
    time.sleep(50)

    # Tentar se conectar ao Chrome via requests HTTP e injectar um script
    logging.info("Conectando ao Chrome para capturar dados...")

    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By

        # Criar driver
        options = webdriver.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver = webdriver.Chrome(options=options)
        logging.info(f"Conectado! URL: {driver.current_url}")

        # Aguardar mais um pouco e capturar o HTML
        time.sleep(5)

        # Capturar o página source
        page_html = driver.page_source

        # Executar JavaScript para extrair informações dos popups
        js_result = driver.execute_script("""
        return {
            url: window.location.href,
            title: document.title,
            swal_containers: Array.from(document.querySelectorAll('.swal2-container')).map(el => ({
                visible: !el.style.display.includes('none'),
                innerHTML: el.innerHTML.substring(0, 500),
                classes: el.className,
            })),
            modals: Array.from(document.querySelectorAll('[role="dialog"], .modal')).map(el => ({
                visible: el.offsetParent !== null,
                text: el.textContent.substring(0, 100),
                tag: el.tagName,
            })),
            buttons_visible: Array.from(document.querySelectorAll('button')).filter(b => {
                const style = window.getComputedStyle(b);
                return style.display !== 'none' && b.offsetParent !== null;
            }).map(b => ({
                text: b.textContent.trim().substring(0, 30),
                class: b.className,
                id: b.id,
                tagName: b.tagName
            }))
        };
        """)

        # Salvar os dados
        output = {
            'timestamp': datetime.now().isoformat(),
            'url': driver.current_url,
            'page_source_length': len(page_html),
            'javascript_result': js_result,
            'page_source': page_html  # Salvar full HTML
        }

        output_file = 'C:\\workspace\\SOMA\\popup_capture.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2, ensure_ascii=False)

        logging.info(f"✓ Dados capturados e salvos em: {output_file}")
        logging.info(f"\nBotões visíveis encontrados:")
        for btn in js_result.get('buttons_visible', []):
            logging.info(f"  - {btn}")

        logging.info(f"\nSweetAlert containers encontrados:")
        for swal in js_result.get('swal_containers', []):
            logging.info(f"  - Visível: {swal['visible']}, Classes: {swal['classes']}")

        # Tirar screenshot
        driver.save_screenshot('C:\\workspace\\SOMA\\artifacts\\screenshots\\popup_capture.png')
        logging.info("✓ Screenshot capturado")

        driver.quit()

    except Exception as e:
        logging.error(f"Erro ao capturar dados: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Finalizar SOMA
        logging.info("Finalizando SOMA...")
        soma_proc.terminate()
        try:
            soma_proc.wait(timeout=10)
        except:
            soma_proc.kill()
        logging.info("Concluído")

if __name__ == '__main__':
    main()
