import sys
sys.path.insert(0, 'C:\\workspace\\SOMA\\src')

import time
import os
from pathlib import Path

# Iniciar SOMA em background com HEADLESS=false
os.environ['HEADLESS'] = 'false'

# Importar após definir variável
from soma_app.workflows.run_soma import main

print("Iniciando SOMA em modo visível...")
print("O Chrome abrirá em segundos...")
print("")
print("Enquanto SOMA executa:")
print("  1. Deixa a página carregar completamente")
print("  2. Clica em 'Sim' quando o popup aparecer")
print("  3. O script vai tirar screenshots automaticamente")
print("  4. Os screenshots aparecerão em artifacts/screenshots/")
print("")

try:
    main()
except KeyboardInterrupt:
    print("\nSOMА interrompido pelo utilizador")
except Exception as e:
    print(f"\nErro ao executar SOMA: {e}")
