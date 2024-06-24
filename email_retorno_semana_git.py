import yfinance as yf
import pandas as pd
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta


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
    df_ohlc=yf.download(l_ticker_yf,actions=True,start=comeco,end=final,interval='1wk',auto_adjust=True)
    df_ohlc=df_ohlc.dropna()#remover valores faltantes
    return_w=(df_ohlc['Close']-df_ohlc['Open'])/df_ohlc['Open']*100#calculando retorno da semana
    return_w=return_w.T.sort_values(by=return_w.index.strftime('%Y-%m-%d')[0],ascending=False)#transpondo o df e ordenando do maior retorno para o menor
    return_w.index=return_w.index.str.replace('.SA', '')
    return_w.columns = ['retorno']
    return_w.index.name = 'Código de Negociação'
    df_tot=pd.merge(b3, return_w, left_on="Código de Negociação", right_index=True).sort_values(by='retorno',ascending=False)
    close=df_ohlc['Close'].T
    close.index=close.index.str.replace('.SA', '')
    close.columns = ['Close']
    close.index.name = 'Código de Negociação'
    df_tot=pd.merge(df_tot, close, left_on="Código de Negociação", right_index=True).sort_values(by='retorno',ascending=False)
    divid=df_ohlc['Dividends'].T
    divid.index=close.index.str.replace('.SA', '')
    divid.columns = ['Dividends']
    divid.index.name = 'Código de Negociação'
    df_tot=pd.merge(df_tot, divid, left_on="Código de Negociação", right_index=True).sort_values(by='retorno',ascending=False)
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
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
        elif df.loc[i]['retorno']<0:
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
        else:
          txt_email_maior+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'<td>'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
    txt_email_maior+='</table></body>'

    txt_email_menor=cria_tabela_html('Maiores baixas da semana',['Código de Negociação','Tipo','Retorno'])
    for i in df.tail(3).iloc[::-1].index:
        if df.loc[i]['retorno']>0:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
        elif df.loc[i]['retorno']<0:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
        else:
          txt_email_menor+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
    txt_email_menor+='</table></body>'
    txt_email+=txt_email_maior+txt_email_menor
    return txt_email


def cria_tabela_carteira(df,style):
    txt_email=style
    txt_email_carteira=cria_tabela_html('Retornos Carteira',['Código de Negociação','Tipo','Preço','Retorno'])
    df=df[['retorno','Tipo','Close']]
    df=df.drop_duplicates()
    for i in df.index:
      if df.loc[i]['retorno']>0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+'R$ '+str(round(df.loc[i]['Close'],2))+'</td>'+'<td style="color: green;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
      elif df.loc[i]['retorno']<0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+'R$ '+str(round(df.loc[i]['Close'],2))+'</td>'+'<td style="color: red;">'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
      else:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+'R$ '+str(round(df.loc[i]['Close'],2))+'</td>'+'<td>'+str(round(df.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
    txt_email_carteira+='</table></body>'

    txt_email+=txt_email_carteira
    return txt_email

def cria_tabela_resumo(df,style,comeco,final):
    txt_email=style
    txt_email_carteira=cria_tabela_html('Resumo Carteira X Indices',['Ativo','Retorno'])
    df=df[['retorno','Quantidade','Close']]
    df.loc[:, 'valor'] = df['Quantidade'] * df['Close']
    df.loc[:, 'peso'] = df['valor'] / df['valor'].sum()
    df.loc[:, 'retorno_ponderado'] = df['retorno'] * df['peso']
    df_ohlc=yf.download(['^BVSP','^DJI'],start=comeco,end=final,interval='1wk',auto_adjust=True)#pegando os dados do yf
    df_ohlc=df_ohlc.dropna()#remover valores faltantes
    return_w=(df_ohlc['Close']-df_ohlc['Open'])/df_ohlc['Open']*100#calculando retorno da semana
    return_w=return_w.T#transpondo o df e ordenando do maior retorno para o menor
    return_w.index=return_w.index.str.replace('.SA', '')
    return_w.columns = ['retorno']
    return_w.index.name = 'Indice'
    return_w.loc['CARTEIRA']= df['retorno_ponderado'].sum()
    return_w=return_w.sort_values(by='retorno',ascending=False)

    for i in return_w.index:
      if return_w.loc[i]['retorno']>0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td style="color: green;">'+str(round(return_w.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
      elif return_w.loc[i]['retorno']<0:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td style="color: red;">'+str(round(return_w.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
      else:
        txt_email_carteira+='<tr><td>'+i+'</td>'+'<td>'+str(round(return_w.loc[i]['retorno'],2))+'%'+'</td></tr>\n'
    txt_email_carteira+='</table></body>'

    txt_email+=txt_email_carteira

    return txt_email
   



def cria_cabecalho(comeco,final):
    txt_email=f'<h1>{comeco} a {final}</h1>'
    return txt_email



def cria_tabela_dividendos(df,style):
    txt_email=style
    txt_email_carteira=cria_tabela_html('Dividendos',['Código de Negociação','Tipo','Dividendo','DY'])
    df=df[['Código de Negociação','Tipo','Dividends','Close']]
    df=df.drop_duplicates()
    df=df[df['Dividends'] > 0]
    for i in df.index:
      txt_email_carteira+='<tr><td>'+df.loc[i]['Código de Negociação']+'</td>'+'<td>'+df.loc[i]['Tipo']+'</td>'+'<td>'+'R$ '+str(df.loc[i]['Dividends'])+'</td>'+'<td>'+str(round(df.loc[i]['Dividends']/df.loc[i]['Close']*100,2))+'%'+'</td>'+'</tr>\n'
    txt_email_carteira+='</table></body>'

    txt_email+=txt_email_carteira
    return txt_email




final = datetime.now().date() + timedelta(days=-1) # Programa rodando no sábado, pega o dia de sexta-feira
comeco = final + timedelta(days=-4) # Programa rodando no sábado, pega o dia de segunda-feira

caminho_B3=r''#Caminho arquivo carteira da area do investidor B3 .xlsx

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


de=''#Email que vai mandar
para=''#Email que vai receber
txt=''

retorno=retorno_ativos_carteira(caminho_B3,str(comeco),str(final))

txt+=cria_cabecalho(comeco,final)
txt+=cria_tabela_dividendos(retorno,style)
txt+=cria_tabela_resumo(retorno,style,str(comeco),str(final))
txt+=cria_tabela_email_maior_menor(retorno,style)
txt+=cria_tabela_carteira(retorno,style)

with open(r'') as f:#Arquivo com a senha para apps do email .txt
    senha=f.readlines()
    f.close

senha_do_email=senha[0]
enviar_email(de,para,senha_do_email,txt)
