import win32com.client as win32
from pathlib import Path
import datetime, random
import os  
import time
from datetime import date
import re


#### Importando pastas
import lerplanilha
import logs




def lendoEmails():
    destino = Path('C:\\Users\\vinicius.martins\\OneDrive - MSX International Limited\\Documents\\EmailBMW') #"criando uma pasta. the create file"
    #destino.mkdir(parents=True, exist_ok=True)
    file_log = logs.criarLogs()
    print("Log criado")
    ##Criando Msg de erro para controle do LOG
    msg_logs = 'Lendo Email'
    logs.escrevendoLog(file_log,msg_logs)
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    imbox = outlook.GetDefaultFolder(6) #Selecionar a caixa de entrada (6)
    #pegar uma pasta especifica 
    #imbox = rootFolder.Folders["Caixa de Entrada"].Folders["importantes"].Folders["testes"]
    messages = imbox.items.Restrict("[Unread]=true")
    msg = messages.GetLast()
    #colocar o modelo de subject
    #teste = messages.Restrict("[Subject] = 'teste'")
    
    destExcel = ''
    for msg in messages:
        print("Lendo Emails...")
        definirMSG = 0
        subject = msg.Subject
        msg.Unread = False
        print("ler email =  ", subject )
        if msg.Subject == "VENDA SI - SOLICITAÇÃO DE REGISTRO" or msg.Subject == "URGENTE - VENDA SI - SOLICITAÇÃO DE REGISTRO":
            
            ###ESCREVENDO LOG###
            msg_logs = 'Email encontrado'
            logs.escrevendoLog(file_log,msg_logs)
            
            body = msg.body
            anexo = msg.Attachments 
            print("Tratando email:   ", subject )
            #responderEmails(msg,definirMSG)
            if len(anexo) >= 2:
                #extrair    nome_criacao_pasta chassi
                nome_criacao_pasta = extrairBody(body)
                print(nome_criacao_pasta)
                #crianda    nome_criacao_pasta da pasta do anexo com o subject 
                
                test =  'teste'
                nome_criacao_pasta =  nome_criacao_pasta.replace("/","").replace(" ","-")
                #print(len  nome_criacao_pasta))
                pastaDestino = os.path.abspath(os.path.join(destino, nome_criacao_pasta))
                print("Criando pasta",pastaDestino )
                #criando pasta
                
                try:
                    if not os.path.exists(pastaDestino):
                        # abspath is just to simplify out the path so error messages are plainer
                        # while os.path.join ensures the path is constructed with OS preferred separators
                        os.makedirs(pastaDestino , mode = 0o666)
                        print("criando")
                    dest = pastaDestino 
                    print("Criando pastas com o Subject ",subject )
                    i = 1
                    for att in anexo:
                        print("Anexos Identificado" , att)
                        destArquivo = os.path.join(pastaDestino,str(att))
                        att.SaveAsFile(destArquivo)
                        if ".xlsm" in str(att) or ".xlsx" in str(att) or ".csv" in str(att):
                            print("Excel Identificado") 
                            #ler planilha anexada
                            lerplanilha.mainPlanilha(destArquivo)
                            definirMSG=1
                            
                            #responderEmails(msg,definirMSG)
                            msg_logs = 'Preencheu arquivo'
                            logs.escrevendoLog(file_log,msg_logs)
                            
                        i +=1
                        print("Preenchimento arquivo Concluido")
                        ###ESCREVENDO LOG###

                        
                    ###ESCREVENDO LOG###    
                    msg_logs = 'Execução Concluida'
                    logs.escrevendoLog(file_log,msg_logs)
                    
                         
                except OSError as err:
                    print("OS error:", err)
                except ValueError:
                    print("Could not convert data to an integer.")
                except Exception as err:
                    print(f"Unexpected {err=}, {type(err)=}")
                        ###ESCREVENDO LOG###
                    msg_logs = 'Erro no Preenchimento' + str(err)
                    logs.escrevendoLog(file_log,msg_logs)
                    raise
                    print('erro')
                    definirMSG = 2
                    #responderEmails(msg,definirMSG )
            else:
                ###ESCREVENDO LOG###
                msg_logs = 'Anexos Faltantes'
                logs.escrevendoLog(file_log,msg_logs)
                
                definirMSG = 2
                #responderEmails(msg,definirMSG )
                #responderEmails(email,definirMSG)
                time.sleep(2)
                 
                
    return destExcel
#função para responder os emails, a variavel definirResposta define um retorno para o usuario de acordo com dados numericos



def mensagemResposta(definirResposta):
    # obtém a hora atual para bom dia, boa tarde ou boa noite
    hora_atual = datetime.datetime.now().hour

    if hora_atual > 12:
        mensagem = 'Bom dia!'
    elif 12 <= hora_atual < 18:
        mensagem = 'Boa tarde!'
    else:
        mensagem = 'Boa noite!'
    if definirResposta == 1:
        #email de resposta dos dados ok
        text1 =  mensagem + " Atenção: esta é uma mensagem automática! Não responda este email! \n\ n Confirmamos o recebimento da sua solicitação de ativação de pacote Service Inclusive.\n Atenciosamente,\n Service Inclusive Brasil"
    elif definirResposta == 2:
        #email de resposta de erro
        text1 =  mensagem + " ***ATENÇÃO: Esta é uma mensagem automática, NÃO RESPONDA ESTE E-MAIL!*** \n \n Prezados,\n \n Identificamos um problema na solicitação que impede a ativação do Pacote Service Inclusive.\n Por favor, verifique as informações e realize um novo envio com os documentos abaixo \n - Formulário de venda em Excel;\n - Termo de adesão; \n - Cópia do Key Reader ou Nota Fiscal. \n \n ***Devido ao problema identificado no envio, informamos que a solicitação enviada NÃO SERÁ PROCESSADA, por favor, corrija seguindo as instruções e nos envie novamente*** \n \n Atenciosamente, \n Service Inclusive Brasil"
    return text1

#Pegar o numero do chassi para nomear a pasta
def extrairBody (nome_criacao_pasta):
    criacao_pasta =  nome_criacao_pasta.lower()
    if  'chassi' in criacao_pasta or  "chassi:" in criacao_pasta:
        result = re.search(r'chassi: (.*)$',   criacao_pasta, re.MULTILINE)
        nome_criacao_pasta = str(result.group(1)).rstrip()
    print("teste")
    return  nome_criacao_pasta
               

def responderEmails(msg,definirMSG ):
    body = mensagemResposta(definirMSG)
    reply = msg.Reply()
    reply.Body = body + reply.Body
    reply.Send()
    '''mail.HTMLBody = '<h2></h2>' #this field is optional'''
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    print("email enviado")


##########################


#Criando logs        



''' Function to check if outlook is open
def check_outlook_open ():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        # Append to the list of process
        list_process.append(p.name())
    # If outlook open then return True
    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False'''

    
      