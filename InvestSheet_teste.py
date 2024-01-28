import pandas as pd
from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
acs = []

mod = open("moderador.txt", "r")

def atualizar():
    mod = open("moderador.txt", "r")
    for i in mod:
        coluna, acao = i.split(";")
        acs.append({
            "col": coluna,
            "ac": acao,
        })

def mostrar():
    print("Sua predefinição de ações:\n")
    for a in acs:
        print(a["col"], a["ac"], sep=", ")

def adicionar():
    ac = input("Você quer adicionar qual ação?\n")
    col = input("\nEm qual coluna está essa ação?\n")
    mod = open("moderador.txt", "a")
    mod.write(f"{col.capitalize().strip()};{ac.strip()}")

def procurar(coluna2, ac2):
    chrome = webdriver.Chrome()
    chrome.get(f'https://statusinvest.com.br/acoes/{ac2}')
    preco = chrome.find_element(
        By.XPATH, '//*[@id="main-2"]/div[2]/div/div[1]/div/div[1]/div/div[1]/strong').text
    roe = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[4]/div/div[1]/div/div/strong').text
    dy = chrome.find_element(
        By.XPATH, '//*[@id="main-2"]/div[2]/div/div[1]/div/div[4]/div/div[1]/strong').text
    div_ebitda = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[2]/div/div[2]/div/div/strong').text
    pl = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[1]/div/div[2]/div/div/strong').text
    pvp = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[1]/div/div[4]/div/div/strong').text
    payout = chrome.find_element(
        By.XPATH, '//*[@id="payout-section"]/div/div/div[1]/div[1]/div[1]/strong').text
    cagrL = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[5]/div/div[2]/div/div/strong').text
    cagrR = chrome.find_element(
        By.XPATH, '//*[@id="indicators-section"]/div[2]/div/div[5]/div/div[1]/div/div/strong').text

    tab2[f'{coluna2}1'] = ac2
    tab2[f'{coluna2}3'] = preco
    tab2[f'{coluna2}4'] = roe
    tab2[f'{coluna2}5'] = dy
    tab2[f'{coluna2}6'] = div_ebitda
    tab2[f'{coluna2}7'] = pl
    tab2[f'{coluna2}8'] = pvp
    tab2[f'{coluna2}9'] = payout
    tab2[f'{coluna2}10'] = cagrL
    tab2[f'{coluna2}11'] = cagrR

def main1():

    ac = input("Qual ação você quer colocar em sua planilha?\n")

    while True:
        coluna = input("Qual a coluna em que você quer colocar a ação?\n")
        if tab2[f'{coluna}3'].value == None:
            break
        else:
            sob = int(input(
                "Esta coluna já está ocupada, você escrever sobre o conteúdo da coluna? (isso o deletará)\n"
                "[1] Sim\n[2] Não\n"
            ))
        if sob == 1:
            break

    procurar(coluna, ac)

def main2():
    while True:
        p1 = int(input("[1] Mostrar predefinição\n[2] Adicionar ação\n[3] Preencher planilha\n[4] Sair\n"))
        if p1 == 1:
            mostrar()
        if p1 == 2:
            adicionar()
            atualizar()
        if p1 == 3:
            for n in acs:
                procurar(n["col"], n["ac"])
        if p1 == 4:
            exit()

while True:
    nome = int(
        input("Qual o nome da planilha?\n[1] TABELA DE AÇÕES\n[2] Outro\n"))

    if nome == 1:
        tabela = load_workbook('TABELA DE AÇÕES.xlsx')

    if nome == 2:
        nome2 = input("Qual o nome da planilha?\n")
        tabela = load_workbook(f'{nome2}.xlsx')

    tab2 = tabela.active

    esc1 = int(input("[1] Adicionar ações manualmente\n[2] Utilizar a predefinição\n"))

    if esc1 == 1:
        main1()
    elif esc1 == 2:
        main2()

    if nome == 1:
        tabela.save('TABELA DE AÇÕES.xlsx')
    else:
        tabela.save(f'{nome2}.xlsx')

    while True:
        t = int(input("Você quer colocar outra ação?\n[1] Sim\n[2] Não\n"))
        if t == 2 or t == 1:
            break
        else:
            continue
    if t == 2:
        break
    
