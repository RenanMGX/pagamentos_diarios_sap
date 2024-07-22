import win32com.client
import os
import pandas as pd
import psutil
import subprocess
import xlwings as xw #type: ignore
import traceback

from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Dict
from time import sleep
from Entities.crenciais import Credential
from getpass import getuser
from typing import Literal

class Preparar:
    def __init__(self, *, date:datetime, arquivo_datas:str, em_massa=True) -> None:
        if not isinstance(date, datetime):
            raise TypeError("Apenas datetime é aceito")
        now: datetime = date.replace(hour=0,minute=0,second=0,microsecond=0)
        
        self.empresas:list = ["*"]
        self.__em_massa:bool = em_massa
        
        dias_execucao: Dict[str,datetime] = {
            "dia_1" : now,
            "dia_2" : now + relativedelta(days=1),
            "dia_3" : now + relativedelta(days=2),
            "dia_4" : now + relativedelta(days=3),
            "dia_5" : now + relativedelta(days=4),
            "dia_6" : now + relativedelta(days=5),
            "dia_7" : now + relativedelta(days=6),
            "dia_8" : now + relativedelta(days=7)
        }
        
        if not os.path.exists(arquivo_datas):
            raise FileNotFoundError(f"{arquivo_datas=} não foi encontrado!")
        if not arquivo_datas.endswith("xlsx"):
            raise Exception(f"{arquivo_datas=} apenas arquivos xlsx")
        
        self._fechar_excel(arquivo_datas)
        self.__arquivo_datas: pd.DataFrame = pd.read_excel(arquivo_datas)
        
        self.__datas: dict = self.montar_datas(dias_execucao)
        
        self.__path_files: str = os.getcwd() + "\\arquivos_para_preparar\\"
        if not os.path.exists(self.path_files):
            os.makedirs(self.path_files)
        
        self.__fornecedores_c_debitos_excel:str = "fornecedores_com_debitos.xlsx"
        self.__fornecedores_c_debitos_txt:str = "lista_fornecedores_c_debitos.txt"
        self.__fornecedores_pgto_T_excel:str = "Lista de Fornecedores.xlsx"
        self.__fornecedores_pgto_T_txt:str = "lista_fornecedores_pgto_T.txt"
        self.__lista_relacionais:str = "lista_relacionais.txt"
        
        self.__session: win32com.client.CDispatch
        
    @property
    def path_files(self):
        return self.__path_files
    
    @property
    def arquivo_datas(self):
        return self.__arquivo_datas 
    
    @property
    def datas(self):
        return self.__datas
    
    @property
    def session(self):
        return self.__session
    
    @property
    def fornecedores_c_debitos_excel(self):
        return self.__fornecedores_c_debitos_excel
    
    @property
    def fornecedores_c_debitos_txt(self):
        return self.__fornecedores_c_debitos_txt
    
    @property
    def fornecedores_pgto_T_excel(self):
        return self.__fornecedores_pgto_T_excel
    
    @property
    def fornecedores_pgto_T_txt(self):
        return self.__fornecedores_pgto_T_txt
    
    @property
    def lista_relacionais(self):
        return self.__lista_relacionais
        
    def montar_datas(self, datas_execucao:Dict[str,datetime]) -> dict:
        datas_para_retorno:dict = {}
        
        datas_nao_permitidas:list = self.__arquivo_datas['Data'].astype(str).tolist()
        
        for key,value in datas_execucao.items():
            if not value.strftime('%Y-%m-%d') in datas_nao_permitidas:
                if not ((value.weekday() == 5) or (value.weekday() == 6)):
                    datas_para_retorno[key] = {
                        "data_atual" : value,
                        "data_sap" : value.strftime('%d.%m.%Y'),
                        "data_sap_bmtu" : value.strftime('%d.%m'),
                        "data_sap_atribuicao" : value.strftime('%d.%m.%Y R'),
                        "data_sap_consumo" : value.strftime('%d.%m.%Y O'),
                        "data_sap_imposto" : value.strftime('%d.%m.%Y J')                        
                    }
                else:
                    print(f"{value.strftime('%d.%m.%Y')=}: final de semana")
            else:
                print(f"{value.strftime('%d.%m.%Y')=}: data não permitida")
        
        return datas_para_retorno
    
    def conectar_sap(self, *, user:str, password:str, ambiente:Literal["S4P", "S4Q"]) -> bool:
        try:
            if not self._verificar_sap_aberto():
                print("abrindo programa SAP")
                subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                sleep(5)
            
            print("conectando ao SAP")
            SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
            application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
            connection: win32com.client.CDispatch = application.OpenConnection(ambiente, True) # type: ignore
            self.__session: win32com.client.CDispatch = connection.Children(0)# type: ignore
            
            print("digitando credenciais")
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user # Usuario
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password # Senha
            self.session.findById("wnd[0]").sendVKey(0)
            
            print("SAP conectado!")
            return True
        except Exception as error:
            raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__class__} -> {error}")


 
    # Extrair da FBL1N os fornecedores com partidas em aberto a DÉBITO    
    def primeiro_extrair_fornecedores_fbl1n(self) -> None:
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        try:
            print("Extrair da FBL1N os fornecedores com partidas em aberto a DÉBITO.")
            #self.session.findById("wnd[0]").maximize ()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
            self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
            self.session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press ()
            for empresa in self.empresas:
                self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = empresa #Empresa
            #self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = self.empresas[1] #Empresa
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            self.session.findById("wnd[0]/usr/radX_OPSEL").select()
            self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
            self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
            self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
            self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
            self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = "" # Entrada Data Partidas em Aberto
            #session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = data_sap # Data Inicial de Vencimento
            #session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = data_sap # Data Final de Vencimento
            self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "EMP_C_DEBIT" # Layout
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            #import pdb; pdb.set_trace()
            
            if (aviso_text:=self.session.findById("wnd[0]/sbar").text) == "Nenhuma partida selecionada (ver texto descritivo)":
                print(f"          {aviso_text}")
                raise Exception(aviso_text)
                    
            #import pdb; pdb.set_trace()
            passar:bool = False
            for child_object in self.session.findById("wnd[0]/usr/").Children:
                if child_object.Text == 'Lista não contém dados':
                    print(f"          {child_object.Text}")
                    passar = True
            if passar:
                raise Exception("Lista não contém dados")
            
            
            for file in os.listdir(self.path_files):
                if file == self.fornecedores_c_debitos_excel:
                    try:
                        os.unlink(self.path_files + file)
                    except PermissionError:
                        self._fechar_excel(file_name=file)
                        os.unlink(self.path_files + file)
                    
            
            if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                raise Exception(error)

            self.session.findById("wnd[0]").sendVKey(16)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.path_files
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.fornecedores_c_debitos_excel
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            self._fechar_excel(file_name=self.fornecedores_c_debitos_excel)
        except Exception as error:
            self.fechar_sap()
            raise Exception(error)
        
    def segundo_preparar_documentos(self, *,caminho_fornecedores_pgto_T:str) -> None:
        print("\nPreparando Documentos\n")
        if (not caminho_fornecedores_pgto_T.endswith("\\")) or (not caminho_fornecedores_pgto_T.endswith("/")):
            caminho_fornecedores_pgto_T += "\\"
        if not os.path.exists(caminho_fornecedores_pgto_T):
            raise NotADirectoryError(f"{caminho_fornecedores_pgto_T=} caminho não encontrado!")
        
        self._fechar_excel(self.path_files + self.fornecedores_c_debitos_excel)
        pd.read_excel(self.path_files + self.fornecedores_c_debitos_excel)['Conta'].to_csv(self.path_files + self.fornecedores_c_debitos_txt, header=False, index=False)

        self._fechar_excel(caminho_fornecedores_pgto_T + self.fornecedores_pgto_T_excel)
        pd.read_excel(caminho_fornecedores_pgto_T + self.fornecedores_pgto_T_excel)['Conta'].to_csv(self.path_files + self.fornecedores_pgto_T_txt, header=False, index=False)
        
        # self._fechar_excel(self.path_files + self.fornecedores_c_debitos_excel)  
        # df_relacionais = pd.read_excel(self.path_files + self.fornecedores_c_debitos_excel)
        # df_relacionais = df_relacionais[df_relacionais['Conta'] >= 1100000]
        # df_relacionais = df_relacionais['Conta']
        # df_relacionais.to_csv(self.path_files + self.lista_relacionais, header=False, index=False)
        
        print("Documentos prontos")
    
    # Preparar os documentos na FBL1N do tipo transferência (T).
    def terceiro_preparar_documentos_tipo_t(self) -> None:
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar os documentos na FBL1N do tipo transferência (T).\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    #self.session.findById("wnd[0]").maximize()
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
                    self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
                    self.session.findById("wnd[2]").sendVKey (4)
                    self.session.findById("wnd[3]/usr/ctxtDY_PATH").text = self.path_files + self.__fornecedores_pgto_T_txt
                    self.session.findById("wnd[3]/usr/ctxtDY_PATH").setFocus()
                    self.session.findById("wnd[3]/usr/ctxtDY_PATH").caretPosition = 151
                    self.session.findById("wnd[3]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()
                    self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = self.path_files + self.__fornecedores_c_debitos_txt
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
                    self.session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 154
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self.session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press() #Abrir seleção multipla de Empresas
                    for empresa in self.empresas:
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = empresa #Empresa
                    #self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = self.empresas[1] #Empresa
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self.session.findById("wnd[0]/usr/radX_OPSEL").select()
                    self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
                    self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = ""
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = value['data_sap'] # Data Inicial de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = value['data_sap'] # Data Final de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "TRANSFER" # Layout
                    self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                    
                    
                    if (aviso_text:=self.session.findById("wnd[0]/sbar").text) == "Nenhuma partida selecionada (ver texto descritivo)":
                        print(f"          {aviso_text}")
                        continue
                    
                    #import pdb; pdb.set_trace()
                    passar:bool = False
                    for child_object in self.session.findById("wnd[0]/usr/").Children:
                        if child_object.Text == 'Lista não contém dados':
                            print(f"          {child_object.Text}")
                            passar = True
                    if passar:
                        continue
                    
                    if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                        raise Exception(error)
                    
                    self.session.findById("wnd[0]").sendVKey(5) # Selecionar todas a partidas
                    self.session.findById("wnd[0]/tbar[1]/btn[45]").press() # Modificação em massa
                    sleep(1)
                    self.session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = value['data_sap_atribuicao'] # Alterar Atribuição para pgto
                    
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey (0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                        
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5) 
    
    # Preparar documentos na FBL1N do tipo Boleto (B) que estejam com o DDA cravado.    
    def quarto_preparar_documentos_tipo_b(self) -> None:
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar documentos na FBL1N do tipo Boleto (B) que estejam com o DDA cravado.\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    #self.session.findById("wnd[0]").maximize ()
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
                    self.session.findById("wnd[0]").sendVKey (0)
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
                    self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
                    self.session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press ()
                    for empresa in self.empresas:
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = empresa #Empresa
                    #self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = self.empresas[1] #Empresa
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press ()
                    self.session.findById("wnd[0]/usr/radX_OPSEL").select ()
                    self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
                    self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = ""
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = value['data_sap'] # Data Inicial de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = value['data_sap'] # Data Final de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "Boletos" # Layout
                    self.session.findById("wnd[0]/tbar[1]/btn[8]").press ()
                
                    if (aviso_text:=self.session.findById("wnd[0]/sbar").text) == "Nenhuma partida selecionada (ver texto descritivo)":
                        print(f"          {aviso_text}")
                        continue
                    
                    #import pdb; pdb.set_trace()
                    passar:bool = False
                    for child_object in self.session.findById("wnd[0]/usr/").Children:
                        if child_object.Text == 'Lista não contém dados':
                            print(f"          {child_object.Text}")
                            passar = True
                    if passar:
                        continue
                    
                    if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                        raise Exception(error)
                
                    self.session.findById("wnd[0]").sendVKey(5) # Selecionar todas a partidas
                    self.session.findById("wnd[0]/tbar[1]/btn[45]").press () # Modificação em massa
                    self.session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = value['data_sap_atribuicao']  # Alterar Atribuição para pgto
                    
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey(0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                        
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5)
        
    # Preparar os documentos na FBL1N do tipo Relacionais.
    def quinto_preparar_documentos_relacionais(self) -> None:
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar os documentos na FBL1N Relacionais.\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    #self.session.findById("wnd[0]").maximize()
                    
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
                    self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
                    self.session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press() #Abrir seleção multipla de Empresas
                    for empresa in self.empresas:
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = empresa #Empresa
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self.session.findById("wnd[0]/usr/radX_OPSEL").select()
                    self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
                    self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = ""
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = value['data_sap'] # Data Inicial de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = value['data_sap'] # Data Final de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "RELACIONAIS" # Layout
                    self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                    
                    
                    if (aviso_text:=self.session.findById("wnd[0]/sbar").text) == "Nenhuma partida selecionada (ver texto descritivo)":
                        print(f"          {aviso_text}")
                        continue
                    
                    #import pdb; pdb.set_trace()
                    passar:bool = False
                    for child_object in self.session.findById("wnd[0]/usr/").Children:
                        if child_object.Text == 'Lista não contém dados':
                            print(f"          {child_object.Text}")
                            passar = True
                    if passar:
                        continue
                    
                    if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                        raise Exception(error)
                    
                    self.session.findById("wnd[0]").sendVKey(5) # Selecionar todas a partidas
                    self.session.findById("wnd[0]/tbar[1]/btn[45]").press() # Modificação em massa
                    sleep(1)
                    
                    self.session.findById("wnd[1]/usr/ctxt*BSEG-ZLSCH").text = "T" # altera a forma de pagamento
                    self.session.findById("wnd[1]/usr/txt*BSEG-XREF3").text = value['data_sap_bmtu'] # Alterar Atribuição para pgto
                    
                    
                    #import pdb;pdb.set_trace()
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey (0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                        
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5) 
         
            
    def fechar_sap(self):
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("fechando SAP!")
        try:
            sleep(1)
            self.session.findById("wnd[0]").close()
            sleep(1)
            try:
                self.session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
            except:
                self.session.findById('wnd[2]/usr/btnSPOP-OPTION1').press()
        except Exception as error:
            print(f"não foi possivel fechar o SAP {type(error)} | {error}")
    
    def _fechar_excel(self, file_name:str, *, timeout:int=15) -> None:
        try:
            if "/" in file_name:
                file_name = file_name.split("/")[-1]
            if "\\" in file_name:
                file_name = file_name.split("\\")[-1]
            for _ in range(timeout):
                for app in xw.apps:
                    for open_file in app.books:
                        if file_name.lower() == open_file.name.lower():
                            open_file.close()
                            if len(xw.apps) <= 0:
                                app.kill()
                            return
                sleep(1)
        except:
            print("não foi possivel encerrar o excel")
        
    def _verificar_sap_aberto(self) -> bool:
        for process in psutil.process_iter(['name']):
            if "saplogon" in process.name().lower():
                return True
        return False    

               
if __name__ == "__main__":
    try:
        crd:dict = Credential('SAP_PRD').load()
        
        bot:Preparar = Preparar(date=datetime.now(), arquivo_datas=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/Datas_Execução.xlsx")#, em_massa=False)
        
        bot.conectar_sap(user=crd['user'], password=crd['password'], ambiente='S4P')
        bot.primeiro_extrair_fornecedores_fbl1n()
        bot.segundo_preparar_documentos(caminho_fornecedores_pgto_T=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/")
        bot.terceiro_preparar_documentos_tipo_t()
        bot.quarto_preparar_documentos_tipo_b()
        bot.quinto_preparar_documentos_relacionais()
        
        bot.fechar_sap()
    except Exception as error:
        path = "logs\\"
        if not os.path.exists("logs"):
            os.makedirs("logs")
        
        file = f"{path}log_error_{datetime.now().strftime('%d%m%Y%H%M%S')}"
        with open(file, 'w', encoding='utf-8')as _file:
            _file.write(str(traceback.format_exc()))
    
        raise error        
        
