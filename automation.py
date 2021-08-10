import pyautogui
import pyperclip

#Exportando dados
import pandas as pd

df = pd.read_excel('Example.xlsx') #Faz a leitura do arquivo excel (xlsx)
faturamento = df['Valor Final'].sum() #Faz a leitura da coluna 'Valor Final' e faz a soma de todos os valores dos produtos
qtde_produtos = df['Quantidade'].sum() #Faz a leitura da coluna 'Quantidade' e soma todos os produtos vendidos

pyautogui.PAUSE = 1 #Delay de 1 segundo para cada ação
pyautogui.press('win')
pyautogui.write('google chrome')
pyautogui.press('enter')
pyautogui.write("mail.google.com")
pyautogui.press('enter')
pyautogui.sleep(2)
pyautogui.click(x=71, y=174)
pyautogui.write('SEU EMAIL AQUI')#Insira o e-mail de destino da mensagem
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório mensal de vendas"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
texto = f"""
Prezados, bom dia
Segue abaixo o Relatório total de vendas mensais

O faturamento foi de: R$ {faturamento:,.2f}
e foram vendidos {qtde_produtos:,} produtos

Att
Lucas Pierre
"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','enter')
