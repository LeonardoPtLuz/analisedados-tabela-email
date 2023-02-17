"""
Tutorial ensinado no canal do youtube Hashtag Programação.
Código recebe arquivo excel, separa em três tabelas: Faturamento, Quantidade, Ticket Médio... E envia por e-mail.
Arquivo excel link: https://drive.google.com/drive/folders/1TUDK-Mk2Vo2ea4j1GtAqxnBDNXchXBb-
"""

import pandas as pd
import win32com.client as win32

tab_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None) #Lê todas as colunas


#Faturamento por loja monstrando apenas as colunas ID Loja e Valor Final.
fat = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(fat)


print('-' * 50)
#Quantidade dos produtos vendidos por cada loja.
quant = tab_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quant)


print('-' * 50)
#Ticket médio do produto em cada loja.
ticket_med = (fat['Valor Final'] / quant['Quantidade']).to_frame() #to_frame() transforma em tabela retirando o dtype: float64 no final da tabela ticket_med
ticket_med = ticket_med.rename(columns={0: 'Ticket Médio'}) #Muda o nome da coluna de 0 para Ticker Médio.
print(ticket_med)


#Envia e-mail com relatório.
outlook = win32.Dispatch('outlook.application') #Conecta o python com o outlook do pc.
mail = outlook.CreateItem(0) #Cria email no outlook.
mail.To = 'To address' #Para quem enviar a mensagem(e-mail).
mail.Subject = 'Message subject' #Assunto do e-mail.
mail.HTMLBody = f""" 
<p>Faturamento:</p>
{fat.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade:</p>
{quant.to_html()}

<p>Ticket Médio:</p>
{ticket_med.to_html(formatters={'Ticket Médio': 'R${:,.2f}'})}
"""  #Mensagem a ser enviada(deve ser formatado em HTML).

mail.Send()
print('E-mail enviado!')
