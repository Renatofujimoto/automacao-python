import pandas as pd
import win32com.client as win32

# IMPORTANDO A BASE DE DADOS EM EXCEL E GUARDANDO EM UMA VARIAVEL

tabelas_vendas = pd.read_excel("Vendas.xlsx")

# VISUALIZANDO A BASE DE DADOS

pd.set_option("display.max_columns", None) # tirando o limitador de colunas
print(tabelas_vendas)

# FATURAMENTO POR LOJA (Pegando as colunas ID Loja e Valor Final e somando) obs os nomes tem que ser iguais a tabela

faturamento = tabelas_vendas [["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento)

# QUANTIDADE DE PRODUTOS VENDIDOS POR LOJA

quantidade = tabelas_vendas[["ID Loja" , "Quantidade"]].groupby ("ID Loja").sum()
print(quantidade)

print("-" * 50)

# TICKET MEDIO POR PRODUTO EM CADA LOJA

ticket_medio = (faturamento["Valor Final"] / quantidade ["Quantidade"]).to_frame() # transformando colunas em tabela
ticket_medio = ticket_medio.rename(columns={0:"Ticket Médio"})
print(ticket_medio)

# ENVIANDO POR EMAIL O RELATORIO


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'renatofujimoto2@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>


<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p> 
{faturamento.to_html(formatters={"Valor Final": "R${:,.2f}".format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada loja:</p>
{ticket_medio.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida estou a disposição</p>

<p>Att,</p>
<strong>Renato Fujimoto</strong>
'''

mail.Send()

print("email enviado")