# Passo a passo do projeto

# 1 Acessar o sistema da empresa

## https://dlp.hashtagtreinamentos.com/python/intensivao/login
# pip install pyautogui 

import pyautogui
import time 
from pynput import mouse

# clicar -> pyautogui.click
# escrever -> pyautogui.write
# apertar uma tecla -> pyautogui.press
# apertar atalho -> pyautogui.hotkey("ctrl C e etc")
# pra rolar a página -> pyautogui.scroll 

# apertar a tecla windows 
pyautogui.press("win")
# digitar o nome do programa (chrome)
time.sleep(1.5)
pyautogui.write("edge")
# apertar enter
pyautogui.press("enter") 
time.sleep(1.5)
# digitar o link
link = "https://dlp.hashtagtreinamentos.com/python/intensivao/login"
pyautogui.write(link)
# apertar enter
pyautogui.press("enter")
# espere mais um pouco antes de acessar o link

# 2 Fazer login
time.sleep(1.5)
pyautogui.click(x=686, y=450)
time.sleep(1.5)
pyautogui.write("morerose02@gmail.com")
pyautogui.press("tab")
time.sleep(1.5)
pyautogui.write("soumuitotophahaha")
pyautogui.press("tab")
time.sleep(1.5)
pyautogui.press("enter")

# 3 Importar a base de dados 
# pip install pandas numpy openpyxl

import pandas 

tabela = pandas.read_excel("C:\\Users\\marye\\Documents\\Python Automacao\\data\\produtos.xlsx")
print(tabela) 

for linha in tabela.index:
# 4 Cadastrar o produto
    time.sleep(1)
    pyautogui.click(x=634, y=324)
    #codigo
    codigo = tabela.loc[linha, "codigo"]
    pyautogui.write(codigo)
    pyautogui.press("tab")
    time.sleep(0.5)
    #marca
    marca = tabela.loc[linha, "marca"]
    pyautogui.write(marca)
    pyautogui.press("tab")
    time.sleep(0.5)
    #tipo
    tipo = tabela.loc[linha, "tipo"]
    pyautogui.write(tipo)
    pyautogui.press("tab")
    time.sleep(0.5)
    #categoria
    categoria = tabela.loc[linha, "categoria"]
    pyautogui.write(str(categoria))
    pyautogui.press("tab")
    time.sleep(0.5)
    #preco_unitario
    preco_unitario = tabela.loc[linha, "preco_unitario"]
    pyautogui.write(str(preco_unitario))
    pyautogui.press("tab")
    time.sleep(0.5)
    #custo
    custo = tabela.loc[linha, "custo"]
    pyautogui.write(str(custo))
    pyautogui.press("tab")
    time.sleep(0.5)
    #obs
    obs = tabela.loc[linha, "obs"]
    if not pandas.isna(obs):
        pyautogui.write(obs)
    # enviar produto
    pyautogui.press("tab")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.scroll(50000)

# 5 Repetir ate acabar a base de dados