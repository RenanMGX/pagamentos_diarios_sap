import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from getpass import getuser
from typing import Literal
import traceback
from time import sleep
import xlwings as xw # type: ignore
from sap import SAPManipulation
from copy import deepcopy
import sys
import re
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    from Entities.log_error import LogError
except:
    from log_error import LogError
try:
    from Entities.rotinas import Rotinas, verificarData, RotinasDB, RotinasPeloPortal, RotinaNotFound
except:
    from rotinas import Rotinas, verificarData, RotinasDB, RotinasPeloPortal, RotinaNotFound
    
try:
    from Entities.process import Processos
except:
    from process import Processos
    
class F110Auto(SAPManipulation):
    def __init__(self, *,date:datetime, user:str, password:str, ambiente:str) -> None:
        '''
        Parametros
        dia_execao (int):  Qtd de dias para a execução: 0 = hoje; 1 = amanhâ; 2 = em 2 dias e assim por diante...
        '''
        self.log_error: LogError = LogError()

        self.__data_atual: datetime = date
        self.__data_sap: str = self.__data_atual.strftime('%d.%m.%Y') # Data separada por pontos
        self.__data_sap_atribuicao: str = self.__data_atual.strftime('%d.%m')# Valor da Atribuição
        self.__data_sap_atribuicao2: str = self.__data_atual.strftime('%d.%m.%Y R') # Valor da Atribuição
        self.__data_sap_atribuicao3: str = self.__data_atual.strftime('%d.%m.%Y O') # Valor da Atribuição com O
        self.__data_sap_atribuicao4: str = self.__data_atual.strftime('%d.%m.%Y J') # Valor da Atribuição com J
        self.__data_sap_atribuicao5: str = self.__data_atual.strftime('%d.%m.%Y I') # Valor da Atribuição com I
        self.__data_proximo_dia: str = (self.__data_atual + relativedelta(days=1)).strftime('%d.%m.%Y') # Data do dia seguinte a programação de PGTO 

        self.caminho_arquivo = f"C:\\Users\\{getuser()}\\Downloads\\"
        self.nome_arquivo = f"Relatorio_SAP_{datetime.now().strftime('%d%m%Y%H%M%S')}.xlsx"
        
        
        SAPManipulation.__init__(self, user=user,password=password,ambiente=ambiente)

    def mostrar_datas(self):
        """mostra todas as datas que serão preenchidas pelo programa
        """
        LogError.informativo(f"\n{'-'*20}Datas{'-'*20}")
        LogError.informativo(f"{self.__data_sap=}")
        LogError.informativo(f"{self.__data_sap_atribuicao=}")
        LogError.informativo(f"{self.__data_sap_atribuicao2=}")
        LogError.informativo(f"{self.__data_sap_atribuicao3=}")
        LogError.informativo(f"{self.__data_sap_atribuicao4=}")
        LogError.informativo(f"{self.__data_proximo_dia=}")
        LogError.informativo(f"{'-'*45}\n")
        # if not verificarData(self.__data_atual, caminho=".TEMP/Datas_Execução.xlsx"):
        #     raise Warning(f"está data não é permitida '{self.__data_sap}'")

    def _verificar_conexao(self) -> bool:
        """verifica se o sap está aberto e salva a conexão nas instancias

        Returns:
            bool: True: caso consiga realizar a conexão
                  False: caso não consiga realizar a Conexão
        """
        try:
            self.SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
            self.application: win32com.client.CDispatch = self.SapGuiAuto.GetScriptingEngine
            self.connection: win32com.client.CDispatch = self.application.Children(0)
            self.session: win32com.client.CDispatch = self.connection.Children(0)
            
            return True
        except Exception as error:
            self.log_error.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())
            return False
    
    def _limpar_cache_sap(self) -> None:
        """limpa todas as instancias de conexão com o SAP
        """
        try:
            del self.SapGuiAuto
            del self.application
            del self.connection
            del self.session
        except:
            pass

    def _conectar(self) -> bool:
        """realiza verificação de conexão com o SAP

        Returns:
            bool: True: conectado; False: não conectado
        """
        if self._verificar_conexao() == False:
            self._limpar_cache_sap()
            LogError.informativo("não foi possivel se conectar ao SAP\n  <django:red>")
            return False
        else:
            LogError.informativo("conexão com o SAP estabelecida\n  <django:green>")
            return True

    def listar(self, campo:str) -> None:
        try:
            cont = 0
            for child_object in self.session.findById(campo).Children:
                print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
                cont += 1
        except:
            pass

    def verificar_status(self, campo:str, texto_verificar:str="Proposta de pagamento criada") -> bool:
        """Verifica se o campo informado está retornando a string informada

        Args:
            campo (str): endereço do campo
            texto_verificar (str, optional): texto a ser verificado no campo. Defaults to "Proposta de pagamento criada".

        Returns:
            bool: Retorna True caso encontre ou False caso não encontre
        """
        for status in campo:
            try:
                if texto_verificar.lower() in  self.session.findById(status).Text.lower():
                    return True
            except:
                return False
        return False

    def buscar_campo(self, campo: str) -> list:
        """busca o campo e retorna uma lista dos endereços encontrados

        Args:
            campo (str): endereço a ser procurado

        Returns:
            list: lista de endereços encontrados
        """
        #print(f"Buscar: {campo}")
        lista: list = []
        for child_object in self.session.findById(campo).Children:
            lista.append(child_object.Id.replace("/app/con[0]/ses[0]/", ""))
        return lista

    @SAPManipulation.start_SAP
    def iniciar(self, processo:Processos, empresas_separada:list=[], fechar_sap_no_final:bool=False, salvar_letra:bool=True) -> None:
        LogError.informativo("iniciando checagem das letras")
        if not isinstance(processo, Processos):
            raise TypeError("apenas objeto do tipo Processos Permitido")
        if not isinstance(empresas_separada, list):
            raise TypeError("apenas Listas")
        #procurar_rotinas = Rotinas(self.__data_atual)
        #rotinas_db = RotinasDB(self.__data_atual, ambiente=self.ambiente) #type: ignore
        
        rotinas_portal = RotinasPeloPortal() #type: ignore
        
        # if not self._conectar():
        #     return
        lista: list
        lista_ralacionais:list
        if not empresas_separada:
            if self._extrair_relatorio():                   

                df = pd.read_excel(self.caminho_arquivo + self.nome_arquivo, dtype={'Conta':'str'}).replace(float('nan'), None) 
                
                df_basic = deepcopy(df['Empresa'])
                lista = df_basic.unique().tolist() 
                lista = [x for x in lista if x is not None]
                
                df_contas:pd.DataFrame|pd.Series = deepcopy(df[['Empresa', 'Conta']])
                
                df_contas = df_contas.replace(float('nan'), '0')
                df_contas["Conta"] = df_contas['Conta'].astype(int)
                df_contas = df_contas[df_contas['Conta'] >= 1100000]
                df_contas = df_contas['Empresa']
                
                lista_ralacionais = df_contas.unique().tolist() 
                lista_ralacionais = [x for x in lista_ralacionais if x is not None]
                
                
                LogError.informativo("relatorio da FBL1N terminado  <django:green>")
            else:
                LogError.informativo("sem relatorio  <django:red>")
                return
        else:
            lista = empresas_separada
            lista_ralacionais = empresas_separada


        #lista: list = ['N000']
        #print(lista)
        #rotinas: list = procurar_rotinas.proxima_rotina()
        LogError.informativo("Iniciando lançamento de pagamantos na F110 ")
        
        LogError.informativo("Iniciando lançamentos dos Pagamentos Boletos  <django:blue>")
        #boletos
        if processo.boleto:
            self._SAP_OP(
                lista_empresas=lista,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao,
                rotina_l=rotinas_portal,
                pagamento = "BMTU",
                banco_pagamento = ["PAGTO_BRADESCO", "PAGTO_ITAU"],
                #banco_pagamento = "PAGTO_BRADESCO"
                #rotina=rotinas["primeira"]
            )

            self._SAP_OP(
                lista_empresas=lista,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao2,
                rotina_l=rotinas_portal,
                pagamento = "BMTU",
                banco_pagamento = ["PAGTO_BRADESCO", "PAGTO_ITAU"],
                #banco_pagamento = "PAGTO_BRADESCO"
            )
        
        
        LogError.informativo("Iniciando lançamentos dos Pagamentos Consumo <django:blue>")
        #pagamento O
        if processo.consumo:
            self._SAP_OP(
                lista_empresas=lista,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao3,
                rotina_l=rotinas_portal,
                pagamento="O",
                banco_pagamento = "BRADESCO_TRIBU"
            )
        
        LogError.informativo("Iniciando lançamentos dos Pagamentos Imposto  <django:blue>")
        #pagamento J
        if processo.imposto:
            self._SAP_OP(
                lista_empresas=lista,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao4,
                rotina_l=rotinas_portal,
                pagamento="J",
                banco_pagamento = "BRADESCO_TRIBU"
            )
           
        LogError.informativo("Iniciando lançamentos dos Pagamentos Darfs  <django:blue>")
        #pagamento I
        if processo.darfs:
            self._SAP_OP(
                lista_empresas=lista,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao5,
                rotina_l=rotinas_portal,
                pagamento="I",
                banco_pagamento = "BRADESCO_TRIBU"
            )
            
        LogError.informativo("Iniciando lançamentos dos Pagamentos Relacionais  <django:blue>")
        #pagamento Relacionais
        if processo.relacionais:
            self._SAP_OP(
                lista_empresas=lista_ralacionais,
                data_sap=self.__data_sap,
                data_proximo_dia=self.__data_proximo_dia,
                data_sap_atribuicao=self.__data_sap_atribuicao,
                rotina_l=rotinas_portal,
                pagamento="BMTU",
                banco_pagamento = ["PAGTO_BRADESCO", "PAGTO_ITAU"],
                #banco_pagamento = "PAGTO_BRADESCO",
                relacionais=True
            )
        
        #import pdb;pdb.set_trace()
            
            
#realiza as rotinas no SAP
    def _SAP_OP(
            self,
            lista_empresas: list,
            data_sap: str,
            data_proximo_dia: str,
            data_sap_atribuicao: str,
            rotina_l: RotinasPeloPortal,
            pagamento:Literal["BMTU", "O", "J", "I"],
            banco_pagamento:str|list,
            relacionais:bool = False
    ) -> None:
        '''
        realiza as rotinas no SAP
        Parameters:
        lista_empresas (list) : Lista das empresas que serão executadas
        '''
        
        if not isinstance(lista_empresas, list):
            LogError.informativo("apenas listas  <django:red>")
            return None
        
        #import pdb;pdb.set_trace()    
        for empresa in lista_empresas:
            try:
                if not self.validar_empresa(empresa):
                    raise Exception(f"Empresa {empresa} não é valida")
                

                self.session.findById("wnd[0]").maximize ()
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nf110"
                self.session.findById("wnd[0]").sendVKey (0)

                passou_inicio:bool = False
                
                try:
                    centro_com_letra = rotina_l.get(date=self.__data_atual, centro=empresa, ambiente=self.ambiente)
                except RotinaNotFound as error:
                    raise error
                except:
                    raise Exception(f"não foi possivel encontrar a rotina para a empresa {empresa}")
                
                for _ in range(10):
                    try:
                        sleep(1)
                        CAMPOS_F110 = self.buscar_campo("wnd[0]/usr/")
                        CAMPOS_F110 = self.buscar_campo("wnd[0]/usr/")

                        #import pdb; pdb.set_trace()

                        self.session.findById(CAMPOS_F110[1]).text = data_sap # Data de Execução *** Modificar ***
                        self.session.findById(CAMPOS_F110[3]).text = centro_com_letra # Identificação 
                        passou_inicio = True
                        break
                    except Exception as error:
                        msg_inicio_error = error
                if not passou_inicio:
                    raise Exception(f"{centro_com_letra} - {msg_inicio_error}") #type: ignore

                CAMPOS_F110_ABAS = self.buscar_campo(CAMPOS_F110[4]) #type: ignore

                self.session.findById(CAMPOS_F110_ABAS[1]).select()
                self.session.findById(CAMPOS_F110_ABAS[1]).select()

                if (texto_aviso1:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    LogError.informativo(f"    Aviso: {empresa} == {texto_aviso1}  <django:yellow>")
                    #raise Exception(texto_aviso1)

                
                
                CAMPOS_F110_PARAMETRO = self.buscar_campo(CAMPOS_F110_ABAS[1])
                CAMPOS_F110_PARAMETRO = self.buscar_campo(CAMPOS_F110_PARAMETRO[0])
                
                try:
                    CAMPOS_F110_PARAMETRO_CONTROLE_PAGAMENTO = self.session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssub/1/2/tblSAPF110VCTRL_FKTTAB/").Children
                except:
                    CAMPOS_F110_PARAMETRO_CONTROLE_PAGAMENTO = self.session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/").Children

                for child_object in CAMPOS_F110_PARAMETRO_CONTROLE_PAGAMENTO:
                    if "[0,0]" in child_object.Id:
                        try:
                            self.session.findById(child_object.Id.replace("/app/con[0]/ses[0]/", "")).text = empresa
                        except:
                            raise Exception(f"o codigo '{centro_com_letra}' já foi utilizado para a empresa {empresa}")
                        
                    elif "[1,0]" in child_object.Id:
                        self.session.findById(child_object.Id.replace("/app/con[0]/ses[0]/", "")).text = pagamento
                    
                    elif "[2,0]" in child_object.Id:
                        self.session.findById(child_object.Id.replace("/app/con[0]/ses[0]/", "")).text = data_proximo_dia
                    


                CAMPOS_F110_PARAMETRO_CONTAS = self.buscar_campo(CAMPOS_F110_PARAMETRO[9])

                self.session.findById(CAMPOS_F110_PARAMETRO_CONTAS[1]).text = "*" #Fornecedor
                self.session.findById(CAMPOS_F110_PARAMETRO_CONTAS[6]).text = "*" #Cliente
                self.session.findById(CAMPOS_F110_ABAS[2]).select()

                if (texto_aviso2:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    LogError.informativo(f"    Aviso: {empresa} == {texto_aviso2}  <django:yellow>")
                    #raise Exception(texto_aviso2)

                CAMPOS_F110_SELECAO = self.buscar_campo(CAMPOS_F110_ABAS[2])
                CAMPOS_F110_SELECAO = self.buscar_campo(CAMPOS_F110_SELECAO[0])
                CAMPOS_F110_SELECAO_CRITE_SEL = self.buscar_campo(CAMPOS_F110_SELECAO[1])

                self.session.findById(CAMPOS_F110_SELECAO_CRITE_SEL[1]).setFocus()
                self.session.findById("wnd[0]").sendVKey(4) #Escolher Atribuição como critério
                #session.findById("wnd[1]/usr/lbl[1,6]").setFocus()
                for x in range(5):
                    try:
                        for child_object in self.session.findById("wnd[1]/usr/").Children:
                            campo_para_data:str
                            if relacionais:
                                campo_para_data = "Chave referência 3"
                            else:
                                campo_para_data = "Atribuição"
                            
                            nome_text: str = child_object.Text
                            if campo_para_data.lower() in nome_text.lower():
                                caminho = child_object.Id.replace("/app/con[0]/ses[0]/", "")
                                self.session.findById(caminho).setFocus()
                        break
                    except:
                        sleep(1)

                self.session.findById("wnd[1]").sendVKey(2)

                self.session.findById(CAMPOS_F110_SELECAO_CRITE_SEL[4]).text = data_sap_atribuicao # data Atribuição 

                #import pdb;pdb.set_trace()
                
                
                self.session.findById(CAMPOS_F110_ABAS[3]).select()

                if (texto_aviso3:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    LogError.informativo(f"    Aviso: {empresa} == {texto_aviso3}  <django:yellow>")
                    #raise Exception(texto_aviso3)

                CAMPOS_F110_LOG = self.buscar_campo(CAMPOS_F110_ABAS[3])
                CAMPOS_F110_LOG = self.buscar_campo(CAMPOS_F110_LOG[0])

                self.session.findById(CAMPOS_F110_LOG[1]).selected = "true"
                self.session.findById(CAMPOS_F110_LOG[2]).selected = "true"
                self.session.findById(CAMPOS_F110_LOG[4]).selected = "true"

                CAMPOS_F110_LOG_CONTAS = self.buscar_campo(CAMPOS_F110_LOG[8])

                self.session.findById(CAMPOS_F110_LOG_CONTAS[0]).text = "*"
                self.session.findById(CAMPOS_F110_LOG_CONTAS[2]).text = "*"

                self.session.findById(CAMPOS_F110_ABAS[4]).select()

                if (texto_aviso4:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    LogError.informativo(f"    Aviso: {empresa} == {texto_aviso4}  <django:yellow>")
                    #raise Exception(texto_aviso4)

                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_ABAS[4])
                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_IMPRESS[0])
                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_IMPRESS[1])
                
                if isinstance(banco_pagamento, list):
                    self.session.findById(CAMPOS_F110_IMPRESS[5]).text = banco_pagamento[0] #Banco
                    if len(banco_pagamento) == 2:
                        self.session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRPTAB/ctxtF110V-VARI2[2,2]").text = banco_pagamento[1] #Banco
                else:
                    self.session.findById(CAMPOS_F110_IMPRESS[5]).text = banco_pagamento #Banco

                self.session.findById("wnd[0]/tbar[0]/btn[11]").press () # Gravar Parâmetros
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press ()
                self.session.findById("wnd[0]/tbar[1]/btn[13]").press ()

                try:
                    for child_object in self.session.findById("wnd[1]/usr/").Children:
                        nome: str = child_object.Text
                        if "Exec.imeditamente".lower() in nome.lower():
                            caminho = child_object.Id.replace("/app/con[0]/ses[0]/", "")
                            self.session.findById(caminho).selected = "true"
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except:
                    CAMPOS_ADVERTENCIA = self.buscar_campo("wnd[2]/usr/")
                    if (texto_advertencia:=self.session.findById(CAMPOS_ADVERTENCIA[1]).Text.lower()) != "":
                        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
                        raise Exception(f"{centro_com_letra} - {texto_advertencia}")
                        #print(f"               {texto_advertencia} {empresa + rotina}")

                CAMPOS_F110 = self.buscar_campo("wnd[0]/usr/")
                CAMPOS_F110_ABAS = self.buscar_campo(CAMPOS_F110[4])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_ABAS[0])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_STATUS[0])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_STATUS[1])

                # for x in range(80):
                #     if self.verificar_status(str(CAMPOS_F110_STATUS)):
                #         break
                #     self.session.findById("wnd[0]").sendVKey(14)
                #     sleep(1)
                    
                limite:int = 80 
                for num in range(limite):
                    find_messege:bool = False
                    for caminho in CAMPOS_F110_STATUS:
                        try:
                            if 'Proposta de pagamento criada' in self.session.findById(caminho).Text:
                                find_messege = True
                        except:
                            pass
                    if find_messege:
                        break
                    
                    self.session.findById("wnd[0]").sendVKey(14)
                    sleep(1)
                    if num > (limite-2):
                        raise TimeoutError(f"{centro_com_letra} - não foi possivel identificar se a Proposta de Pagamento foi criada!")                    
                    
                    

                self.session.findById("wnd[0]/tbar[1]/btn[7]").press()

                CAMPOS_PLANEJAR_PAGAMENTO = self.buscar_campo("wnd[1]/usr/")
                self.session.findById(CAMPOS_PLANEJAR_PAGAMENTO[7]).selected = "true"
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                for num in range(limite):
                    find_messege2:bool = False
                    for caminho in CAMPOS_F110_STATUS:
                        try:
                            if 'Programa de pagamento foi executado' in self.session.findById(caminho).Text:
                                find_messege2 = True
                        except:
                            pass
                    if find_messege2:
                        break
                    
                    self.session.findById("wnd[0]").sendVKey(14)
                    sleep(1)
                    if num > (limite-2):
                        raise TimeoutError(f"{centro_com_letra} - não foi possivel identificar se o Programa de pagamento foi executado")                    
                
                self.session.findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").select()
                
                if self.session.findById("wnd[0]/sbar").text == 'Não existem registros de dados para esta seleção':
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    raise Exception(f"--> ARQUIVO NÃO FOI GERADO <-- 'Não existem registros de dados para esta seleção'")
                
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                
                LogError.informativo(f"    Concluido:     {centro_com_letra}  <django:green>")
                self.log_error.register(tipo="Concluido", descri=str(centro_com_letra), trace="")

            except IndexError as error:
                LogError.informativo(f"    Error: {empresa} == Empresa {empresa} não existe na tabela T001 - {error} <django:red>")
                print()
                self.log_error.register(tipo=type(error), descri=f"Empresa {empresa} não existe na tabela T001 - {error}", trace=traceback.format_exc())
            except Exception as error:
                LogError.informativo(f"    Error: {empresa} == {error}  <django:red>")
                self.log_error.register(tipo=type(error), descri=str(f"    Error: {empresa} == {error}"), trace=traceback.format_exc())

    def _extrair_relatorio(self) -> bool:
        LogError.informativo("extraindo relatarios das empresas na FBL1N")
        
        ###########   INICIO
        self.session.findById("wnd[0]").maximize() # Maximiza
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n" # digita a tranzação
        self.session.findById("wnd[0]").sendVKey(0) # aperta ENTER

        try:# se houver algum matcode aberto ele aperta enter para fechar
            self.session.findById("wnd[1]").sendVKey(2) # aperta ENTER
        except:
            pass

        CAMPOS_FBL1N: list = [x.Id.replace("/app/con[0]/ses[0]/", "") for x in self.session.findById("wnd[0]/usr/").Children] # gera uma lista com todos os endereços dentro da transação

        self.session.findById(CAMPOS_FBL1N[2]).text = ""
        self.session.findById(CAMPOS_FBL1N[4]).text = ""

        self.session.findById(CAMPOS_FBL1N[7]).text = "*" # clica no campo da Empresa

        self.session.findById(CAMPOS_FBL1N[22]).text = "" # limpa a data do campo 'Aberto á data fixada'

        try:
            self.session.findById(CAMPOS_FBL1N[44]).text = self.__data_sap # preenche o valor do dia seguinte no campo 'Vencimento líquido'
        except:
            raise Exception("para executar esse script é necessario habilitar o campo Vencimento liquido na transação 'SU3' utilizando o parametro 'FIT_DUE_DATE_SEL'")

        self.session.findById(CAMPOS_FBL1N[50]).text = "/PATRIMAR" # escreve o nome do layout no campo Layout

        self.session.findById(CAMPOS_FBL1N[39]).selected = "true" # marca a flag do campo 'Operações do Razão Especial'
        self.session.findById(CAMPOS_FBL1N[40]).selected = "true" # marca a flag do campo 'Partida-memo'
        self.session.findById(CAMPOS_FBL1N[42]).selected = "true" # marca a flag do campo 'Partida em débito'

        #import pdb;pdb.set_trace()
        
        self.session.findById("wnd[0]").sendVKey(8) # aperta F8 para executar

        try: # veifica se tem algum formulario para ser exibido caso contrario encerra o roteiro
            if self.session.findById('/app/con[0]/ses[0]/wnd[0]/sbar').Text == "Nenhuma partida selecionada (ver texto descritivo)":
                LogError.informativo("Nenhuma partida selecionada (ver texto descritivo)  <django:red>")
                return False
        except:
            pass
        
        if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
            raise Exception(error)
        
        ####### aba dos relatorios
        self.session.findById("wnd[0]").sendVKey(16) # abre a aba para gerar o arquivo excel
        
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press() # aperta em avançar

        CAMPOS_ARQUIVO: list = [x.Id.replace("/app/con[0]/ses[0]/", "") for x in self.session.findById("wnd[1]/usr/").Children] # gera uma lista com todos os endereços dentro da aba

        self.session.findById(CAMPOS_ARQUIVO[1]).text = self.caminho_arquivo # escreve o caminho onde o arquivo será salvo
        self.session.findById(CAMPOS_ARQUIVO[3]).text = self.nome_arquivo # escreve o nome do caminho

        self.session.findById("wnd[1]/tbar[0]/btn[0]").press() #clica em salvar
        
        self._fechar_excel(self.nome_arquivo)

        return True # retorna 
    
    def _fechar_excel(self, file_name:str, timeout:int=10) -> bool:
        try:
            if "/" in file_name:
                file_name = file_name.split("/")[-1]
            if "\\" in file_name:
                file_name = file_name.split("\\")[-1]
            for _ in range(timeout):
                for app in xw.apps:
                    for open_app in app.books:
                        if open_app.name.lower() == file_name.lower():
                            open_app.close()
                            if len(xw.apps) <= 0:
                                app.kill()
                            return True
                sleep(1)
            return False
        except:
            LogError.informativo("não foi possivel encerrar o excel  <django:red>")
            return False
    
    def test(self):
        LogError.informativo("testando F110.py  <django:yellow>")
    
    def validar_empresa(self, empresa:str):
        empresa = str(empresa)
        return bool(re.match(r"[A-Z]{1}\d{3}", empresa))
    
    
if __name__ == "__main__":
    pass
    # register_erro: LogError = LogError()
    # try:
    #     bot: F110 = F110(int(input("dias: "))) #type: ignore
    #     bot.mostrar_datas()
    #     bot.iniciar(Processos())
    # except Exception as error:
    #     print(f"{type(error)} -> {error}")
    #     error_format:str = traceback.format_exc().replace("\n", "|||")
    #     register_erro.register(tipo=type(error), descri=str(error), trace=error_format)
    
    # input()
