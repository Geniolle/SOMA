# 🤖 Orquestrador SOMA - Automação Financeira

Este projeto é um robô de automação web desenvolvido em Python com Selenium. O seu objetivo é ler dados de uma folha de cálculo no Google Sheets e efetuar automaticamente os registos financeiros (Entradas, Saídas e Transferências) no sistema web SOMA, finalizando com a extração dos saldos atualizados de Caixa.

## ✨ Funcionalidades

- **Orquestração Inteligente:** Lê a aba `CONTAORDEM` no Google Sheets e processa apenas as linhas pendentes (onde o `DOC. SOMA` está vazio).
- **Três Fluxos de Trabalho:** Processa automaticamente Entradas, Saídas (com pagamentos e baixas) e Transferências.
- **Atualização de Caixas:** Extrai os saldos finais (Caixa Diário, Caixa Banco, D. Crianças, Verbo Café, Verbo Shop) e atualiza a aba `GERENCIAR CAIXAS`.
- **Manutenção Zero de WebDriver:** Utiliza o `webdriver-manager` para descarregar e atualizar automaticamente o ChromeDriver em background.
- **Integração Modular:** Após terminar as suas tarefas, aciona automaticamente o motor secundário (`src/soma_app/workflows/run_soma.py`) para processamento do histórico.
- **Server-Ready:** Preparado para rodar em servidores Linux (Oracle) em modo invisível (`Headless`), de forma totalmente silenciosa e sem pop-ups.

---

## 🛠️ Pré-requisitos

Para executar este projeto num ambiente local ou servidor, certifique-se de ter:

1. **Python 3.12+** instalado.
2. **Google Chrome** (ou Chromium) instalado nativamente na máquina/servidor.
3. Conta de Serviço do Google Cloud (para manipulação do Sheets).

---

## 🚀 Instalação e Configuração

**1. Clonar o repositório:**
```bash
git clone <url-do-seu-repositorio>
cd SOMA