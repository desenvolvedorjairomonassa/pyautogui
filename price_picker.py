import subprocess
import pyautogui as pag
import sys
import os
import time
import pandas as pd
import datetime
from datetime import date
import ctypes


profit_path = r"C:\Users\AppData\Roaming\Nelogica\Profit\profitchart.exe"

x = pag.size()[0]/50
y = pag.size()[1]/5

subprocess.call(["cmd", "/c", "start", "/max", profit_path])

os.makedirs('csvs', exist_ok=True)

#tempo para o profit estabelecer suas conexões

time.sleep(5)

#timeframe = '1S'
#pag.prompt("Informe o timeframe desejado (1D, 1S, 15...)")
'''
pag.click(x, y)

pag.typewrite(timeframe, 0.05)

pag.press('enter')
'''
lista_nome=["Data", "Compradora", "Valor", "Quantidade", "Vendedora", "Agressor"]
lista = pd.read_csv('ativos.csv', header=None)

#ready = pag.confirm(f"O programa vai começar. Você selecionou o timeframe {timeframe} para {len(lista)} ativos.")
pag.alert('Não mexa no mouse até terminar' )
for ticker in lista.values[:]: #para realizar algum teste, troque [:] por [:5]
    user32 = ctypes.windll.User32
    time.sleep(5)
    pag.click(x, y)

    pag.typewrite(ticker[0], 0.05)
    
    pag.press('enter')

    #tempo para o profit baixar os dados do ativo

    time.sleep(0.5)

    dataCorrente=date.today()
    dataLimite = datetime.date(2021,7,19)
    retroceder = datetime.timedelta(days=1)
    df_ticker=pd.DataFrame(columns=lista_nome)
    
    #if True:
    while dataCorrente>dataLimite:        
        #selecionar data
        
        if (dataCorrente.weekday() == 5 or dataCorrente.weekday() == 6):
            dataCorrente-=retroceder
            continue
            
        pag.rightClick(x, y)
        #print('botão  direito')
        pag.press('down', _pause = False)
        pag.press('enter')
        pag.press('\t')

        anoIni = str(dataCorrente.year)
        mesIni = str(dataCorrente.month)
        diaIni = str(dataCorrente.day)
        pag.typewrite(diaIni,0.05)
        pag.press('right')
        pag.typewrite(mesIni,0.05)
        pag.press('right')
        pag.typewrite(anoIni,0.05)
        for i in range(0,3):
            pag.press('\t')

        anoFin=anoIni
        mesFin=mesIni
        diaFin=diaIni
        pag.typewrite(diaFin,0.15)
        pag.press('right')
        pag.typewrite(mesFin,0.15)
        pag.press('right')
        pag.typewrite(anoFin,0.15)

        #pag.typewrite('17')
             
        pag.press('enter')
        dataCorrente-=retroceder
        pag.rightClick(x, y)
        for a in range(0,10):
            pag.press('down', _pause= False)

        pag.press('enter')
        #df = pd.read_clipboard(lineterminator='\r',sep='\t',names=lista_nome,index_col=None)
        try:
            df = pd.read_clipboard(sep='\t',dtype='unicode') # engine='python',
        except:
            time.sleep(5)
            df = pd.read_clipboard(sep='\t',dtype='unicode') # engine='python',
            

        if (not df.empty):
            #print(df)
            df=df.drop(columns=['Unnamed: 6'])
            df_ticker = pd.concat([df,df_ticker])
            #print(df_ticker)   
        else:
            #print('erro')            
            continue
    df_ticker['Ticker'] = ticker[0]
    df_ticker.to_csv(f'csvs/{ticker[0]}.csv',index=False)
    


pag.alert('Processo terminado')


