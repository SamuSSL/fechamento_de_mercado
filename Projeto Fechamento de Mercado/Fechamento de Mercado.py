import datetime
import yfinance as yf
import win32com.client as win32
from matplotlib import pyplot as plt


cod_de_negociacao = ['^BVSP', 'BRL=X']


hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)

info_mercado = yf.download(cod_de_negociacao, um_ano_atras, hoje)

info_mercado


info_fechamento = info_mercado['Adj Close']

info_fechamento.columns = ['dolar', 'ibovespa']

info_fechamento = info_fechamento.dropna()

info_fechamento


info_anual = info_fechamento.resample('Y').last()

info_mensal = info_fechamento.resample('M').last()

info_anual


retorno_anual_2023 = info_anual.pct_change().dropna()

retorno_mensal = info_mensal.pct_change().dropna()

retorno_diario = info_fechamento.pct_change().dropna()

retorno_anual_2023


retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual_2023.iloc[-1, 0]
retorno_anual_ibov = retorno_anual_2023.iloc[-1, 1]

retorno_anual_ibov


retorno_diario_dolar = round((retorno_diario_dolar * 100), 2)
retorno_diario_ibov = round((retorno_diario_ibov * 100), 2)

retorno_mensal_dolar = round((retorno_mensal_dolar * 100), 2)
retorno_mensal_ibov = round((retorno_mensal_ibov * 100), 2)

retorno_anual_dolar = round((retorno_anual_dolar * 100), 2)
retorno_anual_ibov = round((retorno_anual_ibov * 100), 2)


info_fechamento.plot(y = 'dolar', use_index = True, legend = False)

plt.title('Dólar')

plt.savefig('dolar.png', dpi = 300)

plt.show()


info_fechamento.plot(y = 'ibovespa', use_index = True, legend = False)

plt.title('Ibovespa')

plt.savefig('ibovespa.png', dpi = 300)

plt.show()


outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)


email.To = 'sslemos.05@gmail.com'
email.Subject = 'Relatório diário do mercado'
email.Body = f'''Segue o relatório diário solicitado:

Bolsa:

O Ibovespa conta com uma rentabilidade de {retorno_anual_ibov}% no ano,
enquanto no mês a mesma é de {retorno_mensal_ibov}%.

Já no último dia útil o fechamento foi de {retorno_diario_ibov}%.

Dólar:

O dólar conta com uma rentabilidade de {retorno_anual_dolar}% no ano,
enquanto no mês a mesma é de {retorno_mensal_dolar}%.

Já no último dia útil o fechamento foi de {retorno_diario_dolar}%.

Att,

Samuel Lemos 

'''

anexo_dolar = r"C:\Users\Samuel\Documents\Projetos Python\Projetos\dolar.png"
anexo_ibov = r"C:\Users\Samuel\Documents\Projetos Python\Projetos\ibovespa.png"

email.Attachments.Add(anexo_dolar)
email.Attachments.Add(anexo_ibov)

email.Send()