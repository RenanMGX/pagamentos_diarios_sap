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
from Entities.dependencies.credenciais import Credential
from getpass import getuser
from typing import Literal
from Entities.log_error import LogError
from functools import wraps
from Entities.dependencies.sap import SAPManipulation
from Entities.dependencies.config import Config
from Entities.dependencies.logs import Logs
from Entities.dependencies.functions import Functions

class Preparar(SAPManipulation):
    def __init__(self, *, date:datetime, arquivo_datas:str, em_massa=True, dias:int=8) -> None:
        """Inicializa a classe Preparar.
        Args:
            date (datetime): A data para a qual os documentos estão sendo preparados.
            arquivo_datas (str): O caminho para o arquivo Excel contendo informações de datas.
            em_massa (bool, opcional): Indicador se a operação é em massa. Padrão é True.
            dias (int, opcional): Número de dias para preparar documentos. Padrão é 8.
        Raises:
            TypeError: Se a data fornecida não for um objeto datetime.
            FileNotFoundError: Se o arquivo arquivo_datas não existir.
            Exception: Se o arquivo arquivo_datas não for um arquivo xlsx.
        """
        if not isinstance(date, datetime):
            raise TypeError("Apenas datetime é aceito")
        now: datetime = date.replace(hour=0,minute=0,second=0,microsecond=0)
        
        self.empresas:list = ["*"]
        self.__em_massa:bool = em_massa
        
        dias_execucao:Dict[str,datetime] = {
            "dia_1" : now,
        }
        
        if dias > 1:
            for dia in range(dias):
                dias_execucao[f'dia_{dia+2}'] = now + relativedelta(days=(dia + 1))
        
        
        
        if not os.path.exists(arquivo_datas):
            raise FileNotFoundError(f"{arquivo_datas=} não foi encontrado!")
        if not arquivo_datas.endswith("xlsx"):
            raise Exception(f"{arquivo_datas=} apenas arquivos xlsx")
        
        Functions.fechar_excel(arquivo_datas)
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
        
        #self.__session: win32com.client.CDispatch
        
        crd:dict = Credential(Config()['credential']['crd']).load()
        super().__init__(user=crd.get("user"), password=crd.get("password"), ambiente=crd.get("ambiente"))        
        
    @property
    def path_files(self):
        return self.__path_files
    
    @property
    def arquivo_datas(self):
        return self.__arquivo_datas 
    
    @property
    def datas(self):
        return self.__datas
    
    # @property
    # def session(self):
    #     return self.__session
    
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
                    print(f"a data selecionada é {value.strftime('%d.%m.%Y')}, e não pode ser executada pois é final de semana")
            else:
                print(f"a data selecionada {value.strftime('%d.%m.%Y')} não permitida")
        
        return datas_para_retorno
        
    # Extrair da FBL1N os fornecedores com partidas em aberto a DÉBITO 
    @SAPManipulation.start_SAP   
    def primeiro_extrair_fornecedores_fbl1n(self) -> None:
        """
        Extrai fornecedores com itens em aberto a débito da transação FBL1N no SAP.
        Este método executa os seguintes passos:
        1. Verifica se a sessão SAP está conectada.
        2. Navega para a transação FBL1N.
        3. Limpa a seleção de fornecedores.
        4. Insere os códigos das empresas da lista `self.empresas`.
        5. Seleciona itens em aberto e várias caixas de seleção.
        6. Define o layout como "EMP_C_DEBIT".
        7. Executa a busca.
        8. Verifica se há mensagens de aviso ou listas de dados vazias.
        9. Exclui qualquer arquivo Excel existente com o mesmo nome.
        10. Baixa os resultados para um arquivo Excel.
        11. Fecha o arquivo Excel.
        Levanta:
            Exception: Se a sessão SAP não estiver conectada.
            Exception: Se nenhum item for selecionado ou a lista não contiver dados.
            Exception: Se houver um erro de memória no SAP.
        """
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        try:
            print("Extrair da FBL1N os fornecedores com partidas em aberto a DÉBITO.")
            #self.session.findById("wnd[0]").maximize ()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
            self.session.findById("wnd[0]").sendVKey(0)
            
            self.session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = ""
            
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
                        Functions.fechar_excel(file)
                        os.unlink(self.path_files + file)
                    
            
            if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                raise Exception(error)

            self.session.findById("wnd[0]").sendVKey(16)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.path_files
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.fornecedores_c_debitos_excel
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            Functions.fechar_excel(self.fornecedores_c_debitos_excel)
        except Exception as error:
            self.fechar_sap()
            raise Exception(error)
    
    @SAPManipulation.start_SAP
    def segundo_preparar_documentos(self, *, caminho_fornecedores_pgto_T: str) -> None:
        """
        Prepara documentos para pagamentos diários, lendo e gerando arquivos de fornecedores a pagar.

        Args:
            caminho_fornecedores_pgto_T (str): Caminho para a pasta que contém o arquivo Excel de fornecedores.

        Raises:
            NotADirectoryError: Se o caminho fornecido não for encontrado.
        """
        print("\nPreparando Documentos\n")
        if (not caminho_fornecedores_pgto_T.endswith("\\")) or (not caminho_fornecedores_pgto_T.endswith("/")):
            caminho_fornecedores_pgto_T += "\\"
        if not os.path.exists(caminho_fornecedores_pgto_T):
            raise NotADirectoryError(f"{caminho_fornecedores_pgto_T=} caminho não encontrado!")
        
        # self._fechar_excel(self.path_files + self.fornecedores_c_debitos_excel)
        # pd.read_excel(self.path_files + self.fornecedores_c_debitos_excel)['Conta'].to_csv(self.path_files + self.fornecedores_c_debitos_txt, header=False, index=False)

        Functions.fechar_excel(caminho_fornecedores_pgto_T + self.fornecedores_pgto_T_excel)
        pd.read_excel(caminho_fornecedores_pgto_T + self.fornecedores_pgto_T_excel)['Conta'].drop_duplicates().to_csv(self.path_files + self.fornecedores_pgto_T_txt, header=False, index=False)
        
        # self._fechar_excel(self.path_files + self.fornecedores_c_debitos_excel)  
        # df_relacionais = pd.read_excel(self.path_files + self.fornecedores_c_debitos_excel)
        # df_relacionais = df_relacionais[df_relacionais['Conta'] >= 1100000]
        # df_relacionais = df_relacionais['Conta']
        # df_relacionais.to_csv(self.path_files + self.lista_relacionais, header=False, index=False)
        
        print("Documentos prontos")
    
    # Preparar os documentos na FBL1N do tipo transferência (T).
    @SAPManipulation.start_SAP
    def terceiro_preparar_documentos_tipo_t(self) -> None:
        """
        Altera em massa os documentos do tipo transferência (T) na FBL1N, ajustando datas e campos de pagamento.
        """
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar os documentos na FBL1N do tipo transferência (T).\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
                    self.session.findById("wnd[0]").sendVKey(0)
                    
                    self.session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = ""
                    
                    self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
                    sleep(.5)
                    
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN017-LOW").text = "T"
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN015-LOW").text = "RE"

                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN010_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN009_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN009_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    
                    sleep(.5)
                    self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
                    
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()
                    self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
                    self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
                    self.session.findById("wnd[2]").sendVKey (4)
                    self.session.findById("wnd[3]/usr/ctxtDY_PATH").text = self.path_files + self.__fornecedores_pgto_T_txt
                    self.session.findById("wnd[3]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()

                    self.session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "*"
                    #self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
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
                    #import pdb;pdb.set_trace()
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey(0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                    #import pdb;pdb.set_trace()
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5) 
    
    # Preparar documentos na FBL1N do tipo Boleto (B) que estejam com o DDA cravado.
    @SAPManipulation.start_SAP
    def quarto_preparar_documentos_tipo_b(self) -> None:
        """
        Altera em massa os documentos do tipo boleto (B) com DDA, configurando datas de vencimento e atributos de pagamento.
        """
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar documentos na FBL1N do tipo Boleto (B) que estejam com o DDA cravado.\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
                    self.session.findById("wnd[0]").sendVKey (0)
                    
                    self.session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = ""

                    
                    self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
                    sleep(.5)
                    
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN017-LOW").text = "B"
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN015-LOW").text = "RE"

                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN010_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN009_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN009_%_APP_%-VALU_PUSH").press()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").setFocus()
                    self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
                    self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    
                    sleep(.5)
                    self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
                    
                    self.session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "*"
                    #import pdb;pdb.set_trace()
                    #self.session.findById("wnd[1]/tbar[0]/btn[8]").press ()
                    self.session.findById("wnd[0]/usr/radX_OPSEL").select ()
                    self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"
                    self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"
                    self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = ""
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = value['data_sap'] # Data Inicial de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = value['data_sap'] # Data Final de Vencimento
                    self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "Boletos" # Layout
                    
                    #import pdb;pdb.set_trace()
                    
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

                    #import pdb;pdb.set_trace()
                    self.session.findById("wnd[0]").sendVKey(5) # Selecionar todas a partidas
                    self.session.findById("wnd[0]/tbar[1]/btn[45]").press () # Modificação em massa
                    self.session.findById("wnd[1]/usr/txt*BSEG-ZUONR").text = value['data_sap_atribuicao']  # Alterar Atribuição para pgto
                    
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey(0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                    #import pdb;pdb.set_trace()    
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5)
        
    # Preparar os documentos na FBL1N do tipo Relacionais.
    @SAPManipulation.start_SAP
    def quinto_preparar_documentos_relacionais(self) -> None:
        """
        Ajusta documentos classificados como relacionais, alterando forma de pagamento e datas relacionadas.
        """
        try:
            self.session
        except AttributeError:
            raise Exception("o sap precisa ser conectado primeiro!")
        
        print("\nPreparar os documentos na FBL1N Relacionais.\n")
        for key,value in self.datas.items():
                print(f"{key} '{value['data_sap']}' -> Executando!")
                try:
                    #self.session.findById("wnd[0]").maximize()
                    for _ in range(5):
                        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"# Abrir FBL1N
                        self.session.findById("wnd[0]").sendVKey(0)# Abrir FBL1N
                        
                        self.session.findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").text = ""

                        self.session.findById("wnd[0]/tbar[1]/btn[16]").press()# Selecionar
                        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN015_%_APP_%-VALU_PUSH").press()#Abrir seleção multipla de Fornecedores
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select()# Selecionar
                        pd.DataFrame(["FG", "FH", "CS", "FE"]).to_clipboard(index=False, header=False)
                        self.session.findById("wnd[1]/tbar[0]/btn[24]").press()# Colar
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press() # Selecionar todos
                        self.session.findById("wnd[0]/tbar[1]/btn[16]").press()# Selecionar
                        
                        self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").press() #Abrir seleção multipla de Fornecedores
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press()
                        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 3
                        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "3"
                        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
                        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1100000"
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()                        
                        
                        #self.session.findById("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").showContextMenu()#Abrir seleção multipla de Fornecedores
                        #self.session.findById("wnd[0]/usr").selectContextMenuItem ("DELACTX") # eliminar seleção de fornecedores
                        self.session.findById("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").press() #Abrir seleção multipla de Empresas
                        for empresa in self.empresas:
                            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = empresa #Empresa
                        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()# OK
                        
                        self.session.findById("wnd[0]/usr/radX_OPSEL").select()# Selecionar partidas abertas
                        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text = ""# Entrada Data Partidas em Aberto
                        #self.session.findById("wnd[0]/usr/radX_AISEL").select()# Selecionar partidas abertas
                        self.session.findById("wnd[0]/usr/chkX_NORM").selected = "true"# Partidas normais
                        self.session.findById("wnd[0]/usr/chkX_SHBV").selected = "true"# Partidas baixadas
                        self.session.findById("wnd[0]/usr/chkX_MERK").selected = "true"# Partidas marcadas
                        self.session.findById("wnd[0]/usr/chkX_APAR").selected = "true"# Partidas a pagar
                        self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-LOW").text = value['data_sap'] # Data Inicial de Vencimento
                        self.session.findById("wnd[0]/usr/ctxtSO_FAEDT-HIGH").text = value['data_sap'] # Data Final de Vencimento
                        self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "RELACIONAIS" # Layout

                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()# OK

                        try:
                            if self.session.findById("wnd[0]/usr/lbl[5,15]").text == "Consult your SAP administrator.":# Se não tiver permissão para acessar a transação
                                continue
                            break
                        except:
                            break
                    
                    
                    if (aviso_text:=self.session.findById("wnd[0]/sbar").text) == "Nenhuma partida selecionada (ver texto descritivo)":# Se não tiver partidas
                        print(f"          {aviso_text}")
                        continue
                    
                    #import pdb; pdb.set_trace()
                    passar:bool = False
                    for child_object in self.session.findById("wnd[0]/usr/").Children:# Se não tiver partidas
                        if child_object.Text == 'Lista não contém dados':# Se não tiver partidas
                            print(f"          {child_object.Text}")# Se não tiver partidas
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
                    
                    if self.__em_massa:
                        self.session.findById("wnd[1]").sendVKey (0) # **************** Executar Modificação em Massa ****************
                    else:
                        self.session.findById("wnd[1]/tbar[0]/btn[12]").press() ### fechar e não executar em massa
                    #import pdb;pdb.set_trace()    
                    print("          Concluido!")
                    
                except Exception as error:
                    print(f"          Error! {type(error)} -> {error}")
                    print(traceback.format_exc())
        sleep(5) 
         
               
if __name__ == "__main__":
    try:
        crd:dict = Credential('SAP_PRD').load()
        
        date = datetime.now()# + relativedelta(days=1)
        #date = datetime(2025,2,6)
                
        bot:Preparar = Preparar(
            date=date,
            arquivo_datas=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/Datas_Execução.xlsx",
            #dias=1 #<----- desativar para produção
            #em_massa=False
        )
        
        bot.segundo_preparar_documentos(caminho_fornecedores_pgto_T=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/")
        bot.terceiro_preparar_documentos_tipo_t()
        bot.quarto_preparar_documentos_tipo_b()
        bot.quinto_preparar_documentos_relacionais()
        
        bot.fechar_sap()
        Logs(name="Preparar Documento para Pagamento Diario").register(status='Concluido', description=f"Automação concluida em {datetime.now() - date}")
    except Exception as error:
        print(traceback.format_exc())
        Logs(name="Preparar Documento para Pagamento Diario").register(status='Error', description=str(error), exception=traceback.format_exc())
