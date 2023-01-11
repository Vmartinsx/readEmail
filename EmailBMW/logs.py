import datetime, random
import os  
import time
from datetime import date


def criarLogs():
    caminhoDiretorio = 'C:\\Users\\vinicius.martins\\OneDrive - MSX International Limited\\Documents\\EmailBMW\\Logs'
    data_atual = date.today()
    hora_atual = datetime.datetime.now().hour
    minuteAtual = datetime.datetime.now().minute
    hora_atual = str(hora_atual) + '-'+ str(minuteAtual)

    pastaDestino = os.path.abspath(os.path.join(caminhoDiretorio,str(data_atual)))

    if not os.path.exists(pastaDestino):
        print("criando t")
        os.makedirs(pastaDestino , mode = 0o666)
    #criando arquivo
    file =pastaDestino+'\\'+str(hora_atual)

    f= open(file+".txt","w+")
    return file

def escrevendoLog(file, msg):
    hora_atual = datetime.datetime.now().hour
    minuteAtual = datetime.datetime.now().minute
    hora_atual = str(hora_atual) + '-'+ str(minuteAtual)
    msgTXT = "Hora da Execução: "+str(hora_atual) + "- Atividade executada:" + msg
    with open(file+".txt", 'a') as arq:
        arq.write(msgTXT)
        arq.write('\n')
        arq.close
    

