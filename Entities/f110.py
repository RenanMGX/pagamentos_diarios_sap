import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from getpass import getuser
import traceback
from time import sleep
import xlwings as xw # type: ignore

try:
    from log_error import LogError
except:
    from Entities.log_error import LogError

try:
    from rotinas import Rotinas, verificarData
except:
    from Entities.rotinas import Rotinas, verificarData



class F110:
    def __init__(self, dia_execucao:datetime) -> None:
        '''
        Parametros
        dia_execao (int):  Qtd de dias para a execução: 0 = hoje; 1 = amanhâ; 2 = em 2 dias e assim por diante...
        '''
        self.log_error: LogError = LogError()
        # if dia_execucao < 0:
        #     raise ValueError("proibido valores negativos")
        # if isinstance(dia_execucao, float):
        #     dia_execucao = dia_execucao
        # elif not isinstance(dia_execucao, int):
        #     raise TypeError("no parametro 'dia_execucao' apenas valores do tipo (int)")
        
        # #Definir Datas
        # self.__data_atual: datetime = (agora:=datetime.now()) + relativedelta(days=dia_execucao)
        # self.__data_sap: str = self.__data_atual.strftime('%d.%m.%Y') # Data separada por pontos
        # self.__data_sap_atribuicao: str = self.__data_atual.strftime('%d.%m')# Valor da Atribuição
        # self.__data_sap_atribuicao2: str = self.__data_atual.strftime('%d.%m.%Y R') # Valor da Atribuição
        # self.__data_proximo_dia: str = (agora + relativedelta(days=(dia_execucao + 1))).strftime('%d.%m.%Y') # Data do dia seguinte a programação de PGTO 

        self.__data_atual: datetime = dia_execucao
        self.__data_sap: str = self.__data_atual.strftime('%d.%m.%Y') # Data separada por pontos
        self.__data_sap_atribuicao: str = self.__data_atual.strftime('%d.%m')# Valor da Atribuição
        self.__data_sap_atribuicao2: str = self.__data_atual.strftime('%d.%m.%Y R') # Valor da Atribuição
        self.__data_proximo_dia: str = (self.__data_atual + relativedelta(days=1)).strftime('%d.%m.%Y') # Data do dia seguinte a programação de PGTO 

        self.caminho_arquivo = f"C:\\Users\\{getuser()}\\Downloads\\"
        self.nome_arquivo = f"Relatorio_SAP_{datetime.now().strftime('%d%m%Y%H%M%S')}.xlsx"

    def mostrar_datas(self):
        """mostra todas as datas que serão preenchidas pelo programa
        """
        print(f"\n{'-'*20}Datas{'-'*20}")
        print(f"self.__data_sap : '{self.__data_sap}'")
        print(f"self.__data_sap_atribuicao : '{self.__data_sap_atribuicao}'")
        print(f"self.__data_sap_atribuicao2 : '{self.__data_sap_atribuicao2}'")
        print(f"self.__data_proximo_dia : '{self.__data_proximo_dia}'")
        print(f"{'-'*45}\n")
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
            print("não foi possivel se conectar ao SAP\n")
            return False
        else:
            print("conexão com o SAP estabelecida\n")
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
            if texto_verificar.lower() in  self.session.findById(status).Text.lower():
                return True
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

    def iniciar(self) -> None:
        procurar_rotinas = Rotinas(self.__data_atual)

        if not self._conectar():
            return
        
        if self._extrair_relatorio():
            for x in range(5):
                app = xw.Book(self.caminho_arquivo + self.nome_arquivo)
                app.close()
                sleep(1)

            df = pd.read_excel(self.caminho_arquivo + self.nome_arquivo).replace(float('nan'), None) 
            df = df['Empresa']# type: ignore

            lista: list = df.unique().tolist() # type: ignore
            lista = [x for x in lista if x is not None]
        else:
            print("sem relatorio")
            return

        #lista: list = ['N000']

        rotinas: list = procurar_rotinas.proxima_rotina()
        self._SAP_OP(
            lista_empresas=lista,
            data_sap=self.__data_sap,
            data_proximo_dia=self.__data_proximo_dia,
            data_sap_atribuicao=self.__data_sap_atribuicao,
            rotina=rotinas[0]
            #rotina=rotinas["primeira"]
        )

        self._SAP_OP(
            lista_empresas=lista,
            data_sap=self.__data_sap,
            data_proximo_dia=self.__data_proximo_dia,
            data_sap_atribuicao=self.__data_sap_atribuicao2,
            rotina=rotinas[1]
        )

#realiza as rotinas no SAP
    def _SAP_OP(
            self,
            lista_empresas: list,
            data_sap: str,
            data_proximo_dia: str,
            data_sap_atribuicao: str,
            rotina: str,
    ) -> None:
        '''
        realiza as rotinas no SAP
        Parameters:
        lista_empresas (list) : Lista das empresas que serão executadas
        '''
        if not isinstance(lista_empresas, list):
            return None
        
        for empresa in lista_empresas:
            try:
                self.session.findById("wnd[0]").maximize ()
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nf110"
                self.session.findById("wnd[0]").sendVKey (0)

                CAMPOS_F110 = self.buscar_campo("wnd[0]/usr/")


                self.session.findById(CAMPOS_F110[1]).text = data_sap # Data de Execução *** Modificar ***
                self.session.findById(CAMPOS_F110[3]).text = empresa + rotina # Identificação 

                CAMPOS_F110_ABAS = self.buscar_campo(CAMPOS_F110[4])

                self.session.findById(CAMPOS_F110_ABAS[1]).select()
                self.session.findById(CAMPOS_F110_ABAS[1]).select()

                if (texto_aviso1:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    print(f"    Aviso: {empresa+rotina} == {texto_aviso1}")
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
                            raise Exception(f"o codigo '{rotina}' já foi utilizado para esta empresa")
                        
                    elif "[1,0]" in child_object.Id:
                        self.session.findById(child_object.Id.replace("/app/con[0]/ses[0]/", "")).text = "BMTU"
                    
                    elif "[2,0]" in child_object.Id:
                        self.session.findById(child_object.Id.replace("/app/con[0]/ses[0]/", "")).text = data_proximo_dia
                    


                CAMPOS_F110_PARAMETRO_CONTAS = self.buscar_campo(CAMPOS_F110_PARAMETRO[9])

                self.session.findById(CAMPOS_F110_PARAMETRO_CONTAS[1]).text = "*" #Fornecedor
                self.session.findById(CAMPOS_F110_PARAMETRO_CONTAS[6]).text = "*" #Cliente
                self.session.findById(CAMPOS_F110_ABAS[2]).select()

                if (texto_aviso2:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    print(f"    Aviso: {empresa+rotina} == {texto_aviso2}")
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
                            nome_text: str = child_object.Text
                            if "Atribuição".lower() in nome_text.lower():
                                caminho = child_object.Id.replace("/app/con[0]/ses[0]/", "")
                                self.session.findById(caminho).setFocus()
                        break
                    except:
                        sleep(1)

                self.session.findById("wnd[1]").sendVKey(2)

                self.session.findById(CAMPOS_F110_SELECAO_CRITE_SEL[4]).text = data_sap_atribuicao # data Atribuição 

                self.session.findById(CAMPOS_F110_ABAS[3]).select()

                if (texto_aviso3:=self.session.findById("wnd[0]/sbar").Text.lower()) != "":
                    print(f"    Aviso: {empresa+rotina} == {texto_aviso3}")
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
                    print(f"    Aviso: {empresa+rotina} == {texto_aviso4}")
                    #raise Exception(texto_aviso4)

                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_ABAS[4])
                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_IMPRESS[0])
                CAMPOS_F110_IMPRESS = self.buscar_campo(CAMPOS_F110_IMPRESS[1])

                self.session.findById(CAMPOS_F110_IMPRESS[5]).text = "PAGTO_BRADESCO" #Banco

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
                        raise Exception(texto_advertencia)
                        #print(f"               {texto_advertencia} {empresa + rotina}")

                CAMPOS_F110 = self.buscar_campo("wnd[0]/usr/")
                CAMPOS_F110_ABAS = self.buscar_campo(CAMPOS_F110[4])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_ABAS[0])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_STATUS[0])
                CAMPOS_F110_STATUS = self.buscar_campo(CAMPOS_F110_STATUS[1])

                for x in range(80):
                    if self.verificar_status(str(CAMPOS_F110_STATUS)):
                        break
                    self.session.findById("wnd[0]").sendVKey(14)
                    sleep(1)

                self.session.findById("wnd[0]/tbar[1]/btn[7]").press()

                CAMPOS_PLANEJAR_PAGAMENTO = self.buscar_campo("wnd[1]/usr/")
                self.session.findById(CAMPOS_PLANEJAR_PAGAMENTO[7]).selected = "true"
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                for x in range(80):
                    if self.verificar_status(str(CAMPOS_F110_STATUS), texto_verificar="Programa de pagamento foi executado"):
                        break
                    self.session.findById("wnd[0]").sendVKey(14)
                    sleep(1)

                print(f"    Concluido:     {empresa+rotina}")

            except IndexError as error:
                print(f"    Error: {empresa+rotina} == Empresa {empresa} não existe na tabela T001")
                print()
                self.log_error.register(tipo=type(error), descri=f"Empresa {empresa} não existe na tabela T001", trace=traceback.format_exc())
            except Exception as error:
                print(f"    Error: {empresa+rotina} == {error}")
                self.log_error.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())

    def _extrair_relatorio(self) -> bool:
        
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

        # self.session.findById(CAMPOS_FBL1N[7]).setFocus() # clica no campo da Empresa
        # self.session.findById("wnd[0]").sendVKey(4) # abre o matcode

        # CAMPO_EMPRESA: list = self.buscar_campo("wnd[1]/usr/") # faz uma busca dentro do Matcode e retorna uma lista do endereços dos itens encontrados

        # self.session.findById(CAMPO_EMPRESA[4]).caretPosition = 1 # seleciona a primeira empresa
        # self.session.findById("wnd[1]").sendVKey(2) # aperta ENTER

        # self.session.findById(CAMPOS_FBL1N[9]).setFocus() # clica no campo 'até' depois do campo Empresa 
        # self.session.findById("wnd[0]").sendVKey(4) # abre o MatCode

        # # cont = 10
        # # while True: # vai rolando o scroll do do matcode até a quantidade de endereços seja menor que 124 e salva em uma Constante chamada ULTIMA_EMPRESA a quantidade exibida na tela
        # #     self.session.findById("wnd[1]/usr").verticalScrollbar.position = cont
        # #     if (ULTIMA_EMPRESA:=len(self.buscar_campo("wnd[1]/usr/"))) < 124:
        # #         print(ULTIMA_EMPRESA)
        # #         break
        # #     print(ULTIMA_EMPRESA)
        # #     cont += 25
        # #     sleep(1)
        # self.session.findById("wnd[1]/usr").verticalScrollbar.position = 78
        # ULTIMA_EMPRESA = len(self.buscar_campo("wnd[1]/usr/"))

        # self.session.findById(CAMPO_EMPRESA[ULTIMA_EMPRESA-1]).caretPosition = 1 # clica na ultima empresa se baseando na constante 'ULTIMA_EMPRESA' e subtrai 1 do valor para selecionar a ultima empresa exibida
        # self.session.findById("wnd[1]").sendVKey(2) # aperta ENTER

        self.session.findById(CAMPOS_FBL1N[22]).text = "" # limpa a data do campo 'Aberto á data fixada'

        try:
            self.session.findById(CAMPOS_FBL1N[44]).text = self.__data_sap # preenche o valor do dia seguinte no campo 'Vencimento líquido'
        except:
            raise Exception("para executar esse script é necessario habilitar o campo Vencimento liquido na transação 'SU3' utilizando o parametro 'FIT_DUE_DATE_SEL'")

        self.session.findById(CAMPOS_FBL1N[50]).text = "/PATRIMAR" # escreve o nome do layout no campo Layout

        self.session.findById(CAMPOS_FBL1N[39]).selected = "true" # marca a flag do campo 'Operações do Razão Especial'
        self.session.findById(CAMPOS_FBL1N[40]).selected = "true" # marca a flag do campo 'Partida-memo'
        self.session.findById(CAMPOS_FBL1N[42]).selected = "true" # marca a flag do campo 'Partida em débito'

        self.session.findById("wnd[0]").sendVKey(8) # aperta F8 para executar

        try: # veifica se tem algum formulario para ser exibido caso contrario encerra o roteiro
            if self.session.findById('/app/con[0]/ses[0]/wnd[0]/sbar').Text == "Nenhuma partida selecionada (ver texto descritivo)":
                print("Nenhuma partida selecionada (ver texto descritivo)")
                return False
        except:
            pass
        
        ####### aba dos relatorios
        self.session.findById("wnd[0]").sendVKey(16) # abre a aba para gerar o arquivo excel
        
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press() # aperta em avançar

        CAMPOS_ARQUIVO: list = [x.Id.replace("/app/con[0]/ses[0]/", "") for x in self.session.findById("wnd[1]/usr/").Children] # gera uma lista com todos os endereços dentro da aba

        self.session.findById(CAMPOS_ARQUIVO[1]).text = self.caminho_arquivo # escreve o caminho onde o arquivo será salvo
        self.session.findById(CAMPOS_ARQUIVO[3]).text = self.nome_arquivo # escreve o nome do caminho

        self.session.findById("wnd[1]/tbar[0]/btn[0]").press() #clica em salvar

        return True # retorna 
    
    def test(self):
        print("testando F110.py")

if __name__ == "__main__":
    register_erro: LogError = LogError()
    try:
        bot: F110 = F110(int(input("dias: "))) #type: ignore
        bot.mostrar_datas()
        bot.iniciar()
    except Exception as error:
        print(f"{type(error)} -> {error}")
        register_erro.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())
    
    input()
