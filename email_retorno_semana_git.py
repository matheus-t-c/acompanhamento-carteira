import yfinance as yf
import pandas as pd
import smtplib
from email.message import EmailMessage


def enviar_email(From,To,senha,html_text):
  corpo_email = html_text

  msg = EmailMessage()
  msg['Subject'] = "Rentabilidade da carteira na semana"
  msg['From'] = From
  msg['To'] = To
  password = senha
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(corpo_email )

  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
  print('Email enviado')


def retorno_ativos_carteira(caminho,comeco,final):
    #Pegando nome das abas da planilha
    sheet_names = pd.ExcelFile(caminho).sheet_names
    dfs = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(caminho, sheet_name=sheet_name)
        df['Tipo'] = sheet_name
        dfs[sheet_name] = df
    
    b3=pd.concat(dfs.values(), ignore_index=True)
    l_ticker=b3['Código de Negociação'].dropna().tolist()#separando lista de ativos
    l_ticker_yf=[i+'.SA' for i in l_ticker]#padronizando o nome para usar no yf
    df_ohlc=yf.download(l_ticker_yf,start=comeco,end=final,interval='1wk',auto_adjust=True)#pegando os dados do yf
    df_ohlc=df_ohlc.dropna()#remover valores faltantes
    return_w=(df_ohlc['Close']-df_ohlc['Open'])/df_ohlc['Open']*100#calculando retorno da semana
    return_w=return_w.T.sort_values(by=comeco,ascending=False)#transpondo o df e ordenando do maior retorno para o menor
    return_w.index=return_w.index.str[:-3]
    return_w.columns = ['retorno']
    return_w.index.name = 'Código de Negociação'
    df_tot=pd.merge(b3, return_w, left_on="Código de Negociação", right_index=True).sort_values(by='retorno',ascending=False)
    return df_tot

def cria_tabela_html(titulo,l_colunas):
    texto=''
    txt_email=f'''<body>
    <h1>{titulo}</h1>
    <table>
      <tr>'''
    for i in l_colunas:
       txt_email+=f'''
        <th>{i}</th>'''
       
    texto=txt_email
    return texto

def cria_tabela_email_maior_menor(df,style):
    txt_email=style
    txt_email_maior=cria_tabela_html('Maiores altas da semana',['Código de Negociação','Tipo','Retorno'])
    df.set_index('Código de Negociação', inplace=True)
    df=df[['retorno','Tipo']]
    df=df.drop_duplicates()
    for i in df.head(3).index:
        if df.loc[i]['retorno']>0:
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
        elif df.loc[i]['retorno']<0:
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
        else:
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td>'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
    txt_email_maior+='</table></body>'

    txt_email_menor=cria_tabela_html('Maiores baixas da semana',['Código de Negociação','Tipo','Retorno'])
    for i in df.tail(3).iloc[::-1].index:
        if df.loc[i]['retorno']>0:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
        elif df.loc[i]['retorno']<0:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
        else:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
    txt_email_menor+='</table></body>'
    txt_email+=txt_email_maior+txt_email_menor
    return txt_email

def cria_tabela_carteira(df,style):
    txt_email=style
    txt_email_carteira=cria_tabela_html('Retornos Carteira',['Código de Negociação','Tipo','Retorno'])
    df=df[['retorno','Tipo']]
    df=df.drop_duplicates()
    for i in df.index:
      if df.loc[i]['retorno']>0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
      elif df.loc[i]['retorno']<0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
      else:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td>'+str(round(df.loc[i]['retorno'],2))+'</td></tr>\n'
    txt_email_carteira+='</table></body>'

    txt_email+=txt_email_carteira
    return txt_email




comeco='2023-08-14'
final='2023-08-18'

caminho_matheus=r'C:\Users\matheustc\Desktop\Python\Programacao\mercado_financeiro\posicao-2023-08-19-22-13-12.xlsx'

style='''<head>
<title>Tabela Bonita</title>
<style>
  body {
    font-family: Arial, sans-serif;
    background-color: #f4f4f4;
    margin: 0;
    padding: 0;
  }
  table {
    width: 80%;
    margin: 20px auto;
    border-collapse: collapse;
    background-color: #fff;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
  }
  th, td {
    padding: 12px 15px;
    text-align: center;
    border-bottom: 1px solid #ddd;
  }
  th {
    background-color: #f2f2f2;
    color: #333;
    font-weight: bold;
  }
  tr:hover {
    background-color: #f5f5f5;
  }
</style>
</head>'''


de='entretenimentomtv@gmail.com'
para='matheustavaresc@gmail.com'
txt=''
retorno=retorno_ativos_carteira(caminho_matheus,comeco,final)
txt+=cria_tabela_email_maior_menor(retorno,style)
txt+=cria_tabela_carteira(retorno,style)
#print(retorno.columns)

with open(r'C:\Users\matheustc\Desktop\Python\Programacao\mercado_financeiro\senha.txt') as f:
    senha=f.readlines()
    f.close

senha_do_email=senha[0]
#enviar_email(de,para,senha_do_email,txt)