# SOMA — OCR seguro de extratos de cartão

Este projeto trata fotografias de extratos, usa a Google Cloud Vision para OCR, reconstrói as colunas e classifica cada movimento como `VÁLIDO` ou `REVISÃO`.

Por segurança financeira, esta versão **não escreve na sheet CARTÃO** e **não corrige valores por aproximação**.

---

## 1. Autenticação e Requisitos (Conta de Serviço)

A autenticação é feita **exclusivamente por Conta de Serviço (Service Account)** com uma chave JSON local. **Não é necessário e não deve ser utilizado** o login interativo por navegador ou o comando `gcloud auth application-default login`.

### Passo a passo de configuração no Google Cloud Platform:

1. **Criar ou selecionar o projeto GCP**:
   - Projeto: `leitura-ficheiros` (ou o ID configurado em `config.yaml`).

2. **Ativar as APIs necessárias**:
   - Google Drive API;
   - Google Cloud Vision API.

3. **Criar a Conta de Serviço e Chave JSON**:
   - Na Consola GCP, aceda a **IAM & Admin** > **Service Accounts**.
   - Crie uma conta de serviço (ex: `soma-cartao-ocr@leitura-ficheiros.iam.gserviceaccount.com`).
   - Crie uma chave no formato **JSON** e descarregue o ficheiro.

4. **Guardar a chave localmente**:
   - Guarde o ficheiro em:
     `credentials/soma-cartao-ocr.json`
   - *(Opcional)* Pode definir o caminho através da variável de ambiente:
     `SOMA_GOOGLE_CREDENTIALS`

5. **Partilhar a pasta do Google Drive**:
   - Abra o Google Drive.
   - Localize a pasta `EXTRATO_CARTÃO_Images` (ID: `1mwJmtDnPOMrlKkJuUL91iCkYxHI0Noll`).
   - Partilhe a pasta com o e-mail da Conta de Serviço atribuindo o acesso de **Leitor**.

---

## 2. Instalação de Dependências

No terminal do projeto, crie o ambiente virtual (se necessário) e instale as dependências:

```bash
python -m venv .venv
source .venv/bin/activate  # No Linux/macOS
# Ou no Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

---

## 3. Execução

### Modo Google Drive (Padrão)

O `config.yaml` já contém o ID da pasta e o nome da imagem. Execute sem argumentos:

```bash
python main.py
```

O programa utiliza a Conta de Serviço para:
1. Localizar a imagem na pasta `EXTRATO_CARTÃO_Images` (ID: `1mwJmtDnPOMrlKkJuUL91iCkYxHI0Noll`).
2. Descarregar temporariamente o ficheiro.
3. Processar a imagem via OpenCV e Google Cloud Vision API.
4. Gerar os relatórios e estatísticas na pasta `output`.

### Modo Imagem Local

Para testar uma imagem local diretamente, sem consultar o Google Drive:

```bash
python main.py "C:\caminho\para\imagem.jpg"
```

Neste modo, a imagem local é processada diretamente, mas a Conta de Serviço continua a ser utilizada para autenticar na Cloud Vision API.

---

## 4. Ficheiros Gerados na pasta `output/`

- `01_contraste.png`: imagem tratada e ampliada sem sombras;
- `02_binaria.png`: imagem binarizada para diagnóstico;
- `03_diagnostico.png`: delimitação visual de cada linha identificada;
- `movimentos.csv`: extrato estruturado em CSV;
- `movimentos.xlsx`: relatório formatado em Excel com status `VÁLIDO`/`REVISÃO`;
- `resultado.json`: dados completos e metadados da execução.

---

## 5. Regras de Validação Financeira

Uma linha é marcada para `REVISÃO` caso ocorra qualquer uma das situações:
- Data de movimento ou data de valor fora do formato `DD/MM` ou de meses não autorizados;
- Descrição em falta ou demasiado curta;
- Débito ou crédito ausente/ambíguo;
- Nível de confiança do OCR abaixo do limite configurado (`minimum_confidence`).
