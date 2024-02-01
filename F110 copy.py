# *************************** Teste Agendar Execução F110 ***************************
import os
import win32com.client
import sys
from datetime import datetime, timedelta, date
from calendar import monthrange
import time
#import win32com.client as win32
import pandas as pd
import json
from getpass import getuser
from tkinter import filedialog

caminho = f"C:\\Users\\{getuser()}\\Arquivos SAP"

if not os.path.exists(caminho):
    os.mkdir(caminho)
try:
    with open("parametros.json", 'r')as arqui:
        parametros = json.load(arqui)
except Exception as error:
    print(error)
    print("arquivo não encontrado!")
    input("digite enter para finalizar: ")
    sys.exit()

try:
    parametros['dias']
    parametros['primeira_rotina']
    parametros['segunda_rotina']
except KeyError:
    input("Arquivo parametros.json está incorreto:  ")
    sys.exit()


#Definir Datas
dia_execucao = parametros['dias'] # Qtd de dias para a execução: 0 = hoje; 1 = amanhâ; 2 = em 2 dias e assim por diante...

Data_Atual = datetime.now() + timedelta(days=dia_execucao, hours=0, minutes=00) # Data de hoje + 1 dia
data_sap = Data_Atual.strftime('%d.%m.%Y') # Data separada por pontos
data_sap_atribuicao = Data_Atual.strftime('%d.%m.%Y') + " R" # Valor da Atribuição
data_sap_atribuicao2 = Data_Atual.strftime('%d.%m')# Valor da Atribuição
Data_Atual_MaisDois = datetime.now() + timedelta(days=dia_execucao + 1, hours=0, minutes=00) # Data do dia seguinte a programação de PGTO 
Data_Atual_MaisDois_Sap = Data_Atual_MaisDois.strftime('%d.%m.%Y') # Formato SAP

#Definir empresas FBL1N
def SAP_OP_1():
    
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:

        application = None
        SapGuiAuto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
    session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press ()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*" #Empresa
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "*" #Empresa
    session.findById("wnd[1]/tbar[0]/btn[8]").press ()

    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN010_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = data_sap_atribuicao
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = data_sap_atribuicao2
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[8]").press()


    session.findById("wnd[0]/usr/radX_OPSEL").select ()
    session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
    session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
    session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
    session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = "" # Entrada Data Partidas em Aberto
    session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = data_sap # Data Inicial de Vencimento
    session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = data_sap # Data Final de Vencimento
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "PGTOS_DIA" # Layout
    session.findById("wnd[0]/tbar[1]/btn[8]").press ()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho #### RENAN ALTERAR
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Lista Empresas.txt" #### RENAN ALTERAR
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 29
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    session = None          
    connection = None
    application = None
    SapGuiAuto = None

#SAP_OP_1()


#Definir empresas e criar uma lista
#empresas = pd.read_csv('dist/empresas.txt', delimiter='\n', skiprows=10, encoding='ISO-8859-1') #### RENAN ALTERAR
#Lista_Empresas = empresas['Empr'].unique().tolist()
print("Selecione o arquivo excel contendo as Empresas:")
caminho = filedialog.askopenfilename()
Lista_Empresas = pd.read_excel(caminho).dropna()['Empr'].unique()
time.sleep(1)
print("Executando F110 com atribuição realizada pelo Robô:")
#Executar F110 com atribuição realizada pelo Robô
for empresa in Lista_Empresas:

    try:

        def SAP_OP_2():
                
                #Logar_SAP.saplogin() #logar no SAP
                
                #excelPath = r'#'
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return

            connection = application.Children(0)
            if not type(connection) == win32com.client.CDispatch:

                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return


            session.findById("wnd[0]").maximize ()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nf110"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxt[0]").text = data_sap # Data de Execução *** Modificar ***
            session.findById("wnd[0]/usr/ctxt[1]").text = empresa + parametros['primeira_rotina'] # Identificação #### RENAN ALTERAR
            session.findById("wnd[0]/usr/ctxt[1]").setFocus ()
            session.findById("wnd[0]/usr/ctxt[1]").caretPosition = (5)
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/txt[0,0]").text = empresa # Empresa
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/ctxt[1,0]").text = "BMTU" # Meio de Pgto
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/ctxt[2,0]").text = Data_Atual_MaisDois_Sap #Data do dia seguinte a programação de PGTO *** Modificar ***
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[0]").text = "*" #Fornecedor
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[1]").text = "*" #Cliente
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[3]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[3]").caretPosition = 1
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/ctxt[0,11]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/ctxt[0,11]").caretPosition = 0
            
            session.findById("wnd[0]").sendVKey (4) #Escolher Atribuição como critério
            #session.findById("wnd[1]/usr/lbl[1,6]").press ()
            session.findById("wnd[1]/usr/lbl[1,6]").setFocus ()
            #session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 7 #Posição da Atribuição
            session.findById("wnd[1]").sendVKey (2)

            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").text = data_sap_atribuicao # Atribuição *** Modificar ***
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").caretPosition = 17
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[0]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[1]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[3]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,0]").text = "*"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,11]").text = "*"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,34]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,34]").caretPosition = 1
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI").select ()

            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssub/1/2/tblSAPF110VCTRL_DRPTAB/ctxt[1,2]").text = "PAGTO_BRADESCO" #Banco

            session.findById("wnd[0]/tbar[0]/btn[11]").press () # Gravar Parâmetros
            session.findById("wnd[0]/tbar[0]/btn[3]").press ()
            session.findById("wnd[0]/tbar[1]/btn[13]").press ()
            session.findById("wnd[1]/usr/chk").selected = "true" #
            session.findById("wnd[1]/usr/chk").setFocus ()
            session.findById("wnd[1]/tbar[0]/btn[0]").press ()

            time.sleep(10)

            session.findById("wnd[0]/tbar[1]/btn[14]").press () # Atualizar
            session.findById("wnd[0]/tbar[1]/btn[21]").press () # Exibir proposta
            session.findById("wnd[0]/tbar[0]/btn[3]").press ()
            session.findById("wnd[0]/tbar[1]/btn[7]").press ()
            session.findById("wnd[1]/usr/chk[1]").selected = "true" #
            session.findById("wnd[1]/usr/chk[1]").setFocus ()
            session.findById("wnd[1]/tbar[0]/btn[0]").press ()

            time.sleep(10)

            session.findById("wnd[0]/tbar[1]/btn[14]").press ()
 
            session = None
            connection = None
            application = None
            SapGuiAuto = None
        SAP_OP_2()
        print(f"    {empresa} --> Concluido")
        
    except:
        print(print(f"    {empresa} --> Error"))
        pass



#Executar F110 com atribuição realizada pela equipe
for empresa in Lista_Empresas:

    try:

        def SAP_OP_3():
                
                #Logar_SAP.saplogin() #logar no SAP
                
                #excelPath = r'#'
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return

            connection = application.Children(0)
            if not type(connection) == win32com.client.CDispatch:

                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return


            session.findById("wnd[0]").maximize ()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nf110"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxt[0]").text = data_sap # Data de Execução *** Modificar ***
            session.findById("wnd[0]/usr/ctxt[1]").text = empresa + parametros['segunda_rotina'] # Identificação #### RENAN ALTERAR
            session.findById("wnd[0]/usr/ctxt[1]").setFocus ()
            session.findById("wnd[0]/usr/ctxt[1]").caretPosition = (5)
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/txt[0,0]").text = empresa # Empresa
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/ctxt[1,0]").text = "BMTU" # Meio de Pgto
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/ctxt[2,0]").text = Data_Atual_MaisDois_Sap #Data do dia seguinte a programação de PGTO *** Modificar ***
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[0]").text = "*" #Fornecedor
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[1]").text = "*" #Cliente
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[3]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/sub/1/2/2/ctxt[3]").caretPosition = 1
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/ctxt[0,11]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/ctxt[0,11]").caretPosition = 0
            
            session.findById("wnd[0]").sendVKey (4) #Escolher Atribuição como critério
            #session.findById("wnd[1]/tbar[0]/btn[0]").press ()
            session.findById("wnd[1]/usr/lbl[1,6]").setFocus ()
            #session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 7 #Posição da Atribuição
            session.findById("wnd[1]").sendVKey (2)

            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").text = data_sap_atribuicao2 # Atribuição *** Modificar ***
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssub/1/2/sub/1/2/1/txt[1,11]").caretPosition = 17
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").select ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[0]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[1]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/chk[3]").selected = "true"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,0]").text = "*"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,11]").text = "*"
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,34]").setFocus ()
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssub/1/2/sub/1/2/1/txt[0,34]").caretPosition = 1
            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI").select ()

            session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssub/1/2/tblSAPF110VCTRL_DRPTAB/ctxt[1,2]").text = "PAGTO_BRADESCO" #Banco

            session.findById("wnd[0]/tbar[0]/btn[11]").press () # Gravar Parâmetros
            session.findById("wnd[0]/tbar[0]/btn[3]").press ()
            session.findById("wnd[0]/tbar[1]/btn[13]").press ()
            session.findById("wnd[1]/usr/chk").selected = "true" #
            session.findById("wnd[1]/usr/chk").setFocus ()
            session.findById("wnd[1]/tbar[0]/btn[0]").press ()

            time.sleep(10)

            session.findById("wnd[0]/tbar[1]/btn[14]").press () # Atualizar
            session.findById("wnd[0]/tbar[1]/btn[21]").press () # Exibir proposta
            session.findById("wnd[0]/tbar[0]/btn[3]").press ()
            session.findById("wnd[0]/tbar[1]/btn[7]").press ()
            session.findById("wnd[1]/usr/chk[1]").selected = "true" #
            session.findById("wnd[1]/usr/chk[1]").setFocus ()
            session.findById("wnd[1]/tbar[0]/btn[0]").press ()

            time.sleep(10)

            session.findById("wnd[0]/tbar[1]/btn[14]").press ()
 
            session = None
            connection = None
            application = None
            SapGuiAuto = None
        SAP_OP_3()
        
    except:
        pass

print("Script Finalizado!")
input()