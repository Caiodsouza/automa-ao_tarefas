import pandas as pd
import pyautogui as pg 
import pyperclip
import time

pg.alert('O programa começara a rodar, por favor não mexer em seu computador enquanto a operação estiver em andamento')
### importar tabela e gerenciar seus dados
tabela = pd.read_excel(r'C:\Users\Lidiane\Desktop\pn_Excel\Vendas - Dez (1).xlsx')
faturamento = tabela ['Valor Final'].sum()
qntd_produtos = tabela ['Quantidade'].sum()
### copiar e enviar dados pelo Gmail
pg.PAUSE = 2
pg.press('WinLeft')
pg.write('Google')
pg.press('Enter')
pyperclip.copy('https://accounts.google.com/signin/v2/identifier?service=mail&passive=1209600&osid=1&continue=https%3A%2F%2Fmail.google.com%2Fmail%2Fu%2F0%2F&followup=https%3A%2F%2Fmail.google.com%2Fmail%2Fu%2F0%2F&emr=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin')
pg.hotkey('Ctrl','v')
pg.press('enter')
time.sleep(2)
pg.click(x=92, y=174)
time.sleep(2)
pg.write('Teste_envio@gmail.com')
pg.click(x=809, y=343)
pg.write('relatorio de vendas')
pg.click(x=899, y=459)
texto = f'''prezados, bom dia

# o faturamento de hoje foi igual a R$: {faturamento:.2f}

e o total de vendas foi de: {qntd_produtos:,}
atenciosamente: Caio Souza'''
pyperclip.copy(texto)
pg.hotkey('ctrl','v')
pg.hotkey('ctrl', 'enter')
pg.press('enter')






