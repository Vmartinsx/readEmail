from datetime import datetime
import pandas as pd
import openpyxl


dest = "C:\\Users\\vinicius.martins\\OneDrive - MSX International Limited\Documents\\EmailBMW\\new Form_Service_Inclusive.xlsm"

#extrai o tamanho do excel
def TamanhoExcel(dest, sheetName ='coleta manual' ):
    dfs = pd.read_excel(dest, sheet_name=sheetName)
    leagthExcel = dfs.shape
    return leagthExcel,dfs

#extrai os dados do excel e verificar se tem dados
def extracaoExcel(leagthExcel, dfs,i):    
    arrows = leagthExcel[0]
    collum = leagthExcel[1]
    print(dfs)
    listDados =  []
    for colluns in range(collum):
        dados = dfs.loc[0][colluns]
        if type(dados) == datetime:
            dados = dados.strftime("%d/%m/%Y")
        else:
            dados = str(dados)
        listDados.append(dados)  
    listDados.insert(0,str(i))
    df = pd.DataFrame([listDados], columns= ['Contrato nº', 'DATA DA VENDA','CONSULTOR','DEPARTAMENTO','CÓDIGO CONCESSIONÁRIA', 'CHASSI', 'KM ATUAL', 'DATA DE INÍCIO GARANTIA', 'NOME COMPLETO DO CLIENTE', 'E-MAIL DO CLIENTE', 'TELEFONE DO CLIENTE', 'CLASSIFICAÇÃO DO VEÍCULO',"MODELO DO VEÍCULO", 'PACOTE CONTRATO (ANOS)', 'PACOTE CONTRATADO (KM)', 'TIPO DE PACOTE', 'PACOTE VÁLIDO ATÉ', 'VALIDADE DO PACOTE EM KM', 'CONCESSIONÁRIA', 'CNPJ', 'CÓDIGO DO PACOTE','VALOR A PAGAR', 'TIPO DE CONTRATO',"VENDA DIRETA BMW SF",'ATIVAÇÃO URGENTE?' ,'PAGO EM ND', 'VALIDAÇÃO','ATIVADO' ] )
    return df
    
    '''listDados =  []
    for arrow in range(0,collum):
        #selecionar a celula
        #tabela modificada pegar a linha, não a coluna 
        dados = dfs.loc[1][arrow]
        if type(dados) == datetime:
            dados = dados.strftime("%d/%m/%Y")
        else:
            dados = str(dados)
        listDados.append(dados)
    #incluir no indice 0 o numero contrato
    i += 1
    listDados.insert(0,str(i))

    #trapor a lista em dataframe. 
    #metodo transpose
    print(len(listDados))
    df = pd.DataFrame([listDados], columns= ['Contrato nº', 'DATA DA VENDA','CONSULTOR','DEPARTAMENTO','CÓDIGO CONCESSIONÁRIA', 'CHASSI', 'KM ATUAL', 'DATA DE INÍCIO GARANTIA', 'NOME COMPLETO DO CLIENTE', 'E-MAIL DO CLIENTE', 'TELEFONE DO CLIENTE', 'CLASSIFICAÇÃO DO VEÍCULO',"MODELO DO VEÍCULO", 'PACOTE CONTRATO (ANOS)', 'PACOTE CONTRATADO (KM)', 'TIPO DE PACOTE', 'PACOTE VÁLIDO ATÉ', 'VALIDADE DO PACOTE EM KM', 'CONCESSIONÁRIA', 'CNPJ', 'CÓDIGO DO PACOTE','VALOR A PAGAR', 'TIPO DE CONTRATO',"VENDA DIRETA BMW SF", 'PAGO EM ND' ], )
    print(df)
    return df
'''
def mainPlanilha(dest):
    
    #path aonde se localiza o excel
    dest1 = "C:\\Users\\vinicius.martins\\OneDrive - MSX International Limited\\Documents\\EmailBMW\\modeloGravarBMW.xlsx"
    
    colunas = ('Contrato nº', 'DATA DA VENDA','CONSULTOR','DEPARTAMENTO','CÓDIGO CONCESSIONÁRIA', 'CHASSI', 'KM ATUAL', 'DATA DE INÍCIO GARANTIA', 'NOME COMPLETO DO CLIENTE', 'E-MAIL DO CLIENTE', 'TELEFONE DO CLIENTE', 'CLASSIFICAÇÃO DO VEÍCULO',"MODELO DO VEÍCULO", 'PACOTE CONTRATO (ANOS)', 'PACOTE CONTRATADO (KM)', 'TIPO DE PACOTE', 'PACOTE VÁLIDO ATÉ', 'VALIDADE DO PACOTE EM KM', 'CONCESSIONÁRIA', 'CNPJ', 'CÓDIGO DO PACOTE','VALOR A PAGAR', 'TIPO DE CONTRATO',"VENDA DIRETA BMW SF", 'PAGO EM ND', 'VALIDAÇÃO', 'ATIVADO' )   
    # only read specific columns from an excel file 
    required_df = pd.read_excel(dest1) 
    print(type(required_df))
    
    
    rows ,arquExcel = TamanhoExcel(dest)
    
    #quantidade de linhas da planilha de inclusão
    row = required_df.shape[0]

    #extract received DF file to email
    df = extracaoExcel(rows,arquExcel,row)
    df3 = pd.DataFrame()
    df3 = pd.concat([required_df,df])
    print(df3.shape)



    writer = pd.ExcelWriter(dest1, engine='xlsxwriter')
    df3.to_excel(writer, sheet_name='welcome', index=False)
    writer.save()
    

mainPlanilha(dest)
