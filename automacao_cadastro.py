"""
AUTOMAÇÃO DE CADASTRO DE PRODUTOS
Autor: Adailson Costa
Descrição:
Script RPA para automatizar o cadastro de produtos em um sistema web
a partir de um arquivo CSV, utilizando PyAutoGUI e Pandas.
"""

# =====================
# IMPORTAÇÕES
# =====================
import pyautogui          # Biblioteca para automação de mouse e teclado
import time               # Biblioteca para controle de tempo/pausas
import pandas as pd       # Biblioteca para leitura e manipulação de dados

# =====================
# CONFIGURAÇÕES GERAIS
# =====================
pyautogui.PAUSE = 1       # Pausa automática de 1 segundo entre ações

# =====================
# ABERTURA DO NAVEGADOR
# =====================
pyautogui.press('win')    # Abre o menu iniciar
pyautogui.write('chrome') # Digita "chrome"
pyautogui.press('enter')  # Abre o navegador

# =====================
# ACESSO AO SITE
# =====================
link = 'https://dlp.hashtagtreinamentos.com/python/intensivao/login'
pyautogui.write(link)
pyautogui.press('enter')

# Aguarda o carregamento da página
time.sleep(3)

# =====================
# LOGIN NO SISTEMA
# =====================
pyautogui.click(x=703, y=377)       # Clica no campo de e-mail
time.sleep(1.5)

pyautogui.write('SEU_EMAIL_AQUI')   # Digite seu e-mail
pyautogui.press('tab')
pyautogui.write('SUA_SENHA_AQUI')   # Digite sua senha
pyautogui.press('tab')
pyautogui.press('enter')

# Aguarda login
time.sleep(3)

# =====================
# LEITURA DO CSV
# =====================
tabela = pd.read_csv('produtos.csv')
print(tabela)

# =====================
# CADASTRO DOS PRODUTOS
# =====================
for linha in tabela.index:

    # Clica no primeiro campo do formulário
    pyautogui.click(x=672, y=298)

    # Código do produto
    pyautogui.write(str(tabela.loc[linha, 'codigo']))
    pyautogui.press('tab')

    # Marca
    pyautogui.write(str(tabela.loc[linha, 'marca']))
    pyautogui.press('tab')

    # Tipo
    pyautogui.write(str(tabela.loc[linha, 'tipo']))
    pyautogui.press('tab')

    # Categoria
    pyautogui.write(str(tabela.loc[linha, 'categoria']))
    pyautogui.press('tab')

    # Preço unitário
    pyautogui.write(str(tabela.loc[linha, 'preco_unitario']))
    pyautogui.press('tab')

    # Custo
    pyautogui.write(str(tabela.loc[linha, 'custo']))
    pyautogui.press('tab')

    # Observações (campo opcional)
    obs = tabela.loc[linha, 'obs']
    if not pd.isna(obs):
        pyautogui.write(str(obs))

    # Envia o formulário
    pyautogui.press('tab')
    pyautogui.press('enter')

    # Scroll para garantir retorno ao topo
    pyautogui.scroll(5000)
