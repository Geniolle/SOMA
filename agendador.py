###################################################################################
# agendador.py
###################################################################################
import time
import subprocess
import sys
import datetime

TEMPO_ESPERA_SEGUNDOS = 120  # 2 minutos

print("🕒 Agendador do Orquestrador SOMA iniciado!")
print(f"⚙️ Configurado para aguardar {TEMPO_ESPERA_SEGUNDOS} segundos entre cada execução completa.")

while True:
    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"\n==================================================================")
    print(f"🚀 Iniciando nova execução do Orquestrador - {agora}")
    print(f"==================================================================")
    
    try:
        # Executa o main.py e ESPERA que ele termine. 
        # Isto garante 100% que nunca haverá execuções paralelas!
        subprocess.run([sys.executable, "main.py"], check=False)
    except Exception as e:
        print(f"❌ Erro crítico ao tentar iniciar o main.py: {e}")
    
    agora_fim = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"\n✅ Execução terminada em {agora_fim}.")
    print(f"⏳ A aguardar {TEMPO_ESPERA_SEGUNDOS} segundos para a próxima ronda...\n")
    
    # Pausa o script durante 2 minutos
    time.sleep(TEMPO_ESPERA_SEGUNDOS)