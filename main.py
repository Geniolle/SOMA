###################################################################################
# main.py (VERSÃO FINAL - SERVER-READY COMPLETAMENTE FUNCIONAL)
###################################################################################

import sys
import os
import subprocess
from pathlib import Path

# FORÇAR O PYTHON A ENXERGAR A PASTA RAIZ
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from selenium_scripts.web_driver import iniciar_webdriver
from selenium_scripts.login import login_soma
from selenium_scripts.processar_saida import processar_saida
from selenium_scripts.processar_entrada import processar_entrada
from selenium_scripts.processar_transferencia import processar_transferencia
from sheets.sheets_service import obter_sheet, obter_todos_os_registros

print("🚀 Iniciando Orquestrador SOMA: Entradas, Transferências e Saídas")

try:
    sheet = obter_sheet()
    registros = obter_todos_os_registros(sheet)
    cabecalho = sheet.row_values(1)
    col_doc_soma = cabecalho.index('DOC. SOMA') + 1 if 'DOC. SOMA' in cabecalho else None

    registros_para_processar = []
    for idx, linha in enumerate(registros, start=2):
        doc_soma = str(linha.get("DOC. SOMA", "")).strip()
        if doc_soma == "" or doc_soma == "Em processamento":
            linha["_row_index"] = idx 
            linha["_status_inicial"] = "Vazio" if doc_soma == "" else "Em processamento"
            registros_para_processar.append(linha)

    # 1. Iniciamos o WebDriver SEMPRE
    driver = iniciar_webdriver()
    login_soma(driver)

    # 2. Processar Linhas (se existirem)
    if not registros_para_processar:
        print("✅ Nenhuma linha pendente para processamento no Excel.")
    else:
        prioridade = {"Entrada": 1, "Transferência": 2, "Saída": 3}
        registros_para_processar.sort(key=lambda x: prioridade.get(str(x.get("TIPO", "")).strip().capitalize(), 99))

        print(f"📊 Total de linhas a processar: {len(registros_para_processar)}")

        for linha in registros_para_processar:
            idx = linha["_row_index"]
            tipo = str(linha.get("TIPO", "")).strip().capitalize()
            status = linha["_status_inicial"]
            
            print(f"\n🔍 Lendo linha {idx} | Tipo: {tipo} | DOC SOMA Atual: [{status}]")
            
            if col_doc_soma and status == "Vazio":
                try:
                    sheet.update_cell(idx, col_doc_soma, "Em processamento")
                except Exception:
                    pass

            try:
                if tipo == "Saída":
                    processar_saida(driver, linha, idx, sheet)
                elif tipo == "Entrada":
                    processar_entrada(driver, linha, idx, sheet)
                elif tipo == "Transferência":
                    # Transferência AGORA ESTÁ OFICIALMENTE LIGADA!
                    processar_transferencia(driver, linha, idx, sheet)
                else:
                    print(f"⚠️ Tipo '{tipo}' desconhecido. A saltar linha {idx}.")
            except Exception as e:
                print(f"❌ Erro na linha {idx}: {e}")

    # 3. Encerrar o driver de forma segura antes de chamar o próximo processo
    driver.quit()
    print("\n✅ Processamento do Orquestrador finalizado com sucesso. WebDriver encerrado.")

    # =========================================================================
    # 4. CHAMAR O PRÓXIMO PROCESSO (run_soma.py)
    # =========================================================================
    print("\n🔗 Passando o testemunho para o run_soma.py...")
    
    # Apontamos diretamente para o caminho correto do seu projeto
    caminho_run_soma = BASE_DIR / "src" / "soma_app" / "workflows" / "run_soma.py"

    if caminho_run_soma.exists():
        try:
            # 1. Copiamos o ambiente atual do Windows
            env = os.environ.copy()
            caminho_src = str(BASE_DIR / "src")
            
            # 2. Ensinamos ao novo terminal onde mora a pasta 'soma_app'
            if "PYTHONPATH" in env:
                env["PYTHONPATH"] = f"{caminho_src};{env['PYTHONPATH']}"
            else:
                env["PYTHONPATH"] = caminho_src

            # 3. Executamos o script com este "novo conhecimento"
            subprocess.run([sys.executable, str(caminho_run_soma)], env=env, check=True)
            print("✅ run_soma.py executado com sucesso e chegou ao fim!")
        except subprocess.CalledProcessError as e:
            print(f"❌ O run_soma.py falhou durante a execução. Erro: {e}")
    else:
        print(f"⚠️ Ficheiro run_soma.py não encontrado em: {caminho_run_soma}")

except Exception as e:
    print(f"❌ Erro fatal no Orquestrador: {e}")
    try:
        driver.quit()
    except:
        pass