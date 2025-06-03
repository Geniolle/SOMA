from selenium import webdriver

def tratar_nomes(nome: str, idade: int, driver: webdriver) -> str: 
    valor_de_retorno = 'O ' + nome + ' tem ' + str(idade) + ' anos.'
    return valor_de_retorno
