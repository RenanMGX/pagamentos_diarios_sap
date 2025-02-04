from Entities.dependencies.logs import Logs
from Entities.dependencies.functions import P
import win32com.client
from functools import wraps
import psutil
import subprocess
from time import sleep
import traceback
import sys

class FindNewID:
    def __init__(self, connection:win32com.client.CDispatch) -> None:
        """
        Inicializa a classe FindNewID com uma lista de IDs de conexão.

        :param connection: Objeto de conexão do SAP.
        """
        self.__connections:list = []
        for x in range(connection.Children.Count):
            self.__connections.append(connection.Children(x).Id)
            
    def target(self, connection:win32com.client.CDispatch):
        """
        Encontra um novo ID de conexão que não está na lista de conexões existentes.

        :param connection: Objeto de conexão do SAP.
        :return: Índice do novo ID de conexão.
        :raises Exception: Se a sessão não for encontrada.
        """
        for x in range(connection.Children.Count):
            if not connection.Children(x).Id in self.__connections:
                return x
        raise Exception("sessão nao encontrada!")

class SAPManipulation():
    @property
    def ambiente(self) -> str|None:
        """
        Retorna o ambiente do SAP.

        :return: Ambiente do SAP.
        """
        return self.__ambiente
    
    @property
    def session(self) -> win32com.client.CDispatch:
        """
        Retorna a sessão atual do SAP.

        :return: Sessão do SAP.
        """
        return self.__session
    @session.deleter
    def session(self):
        """
        Deleta a sessão atual do SAP.
        """
        try:
            del self.__session
        except:
            pass
        
    @property
    def log(self) -> Logs:
        """
        Retorna um objeto de log.

        :return: Objeto de log.
        """
        return Logs()
    
    @property
    def using_active_conection(self) -> bool:
        """
        Retorna se está usando uma conexão ativa.

        :return: True se estiver usando uma conexão ativa, False caso contrário.
        """
        return self.__using_active_conection
    
    def __init__(self, *, user:str|None="", password:str|None="", ambiente:str|None="", using_active_conection:bool=False, new_conection=False) -> None:
        """
        Inicializa a classe SAPManipulation.

        :param user: Usuário do SAP.
        :param password: Senha do SAP.
        :param ambiente: Ambiente do SAP.
        :param using_active_conection: Se está usando uma conexão ativa.
        :param new_conection: Se é uma nova conexão.
        :raises Exception: Se não preencher todos os campos necessários.
        """
        if not using_active_conection:
            if not ((user) and (password) and (ambiente)):
                raise Exception(f"""é necessario preencher todos os campos \n
                                {user=}\n
                                {password=} \n 
                                {ambiente=} \n                            
                                """)
        
        self.__using_active_conection = using_active_conection
        self.__user:str|None = user
        self.__password:str|None = password
        self.__ambiente:str|None = ambiente
        self.__new_connection:bool = new_conection
         
    # Decorador para iniciar o SAP
    @staticmethod
    def start_SAP(f):
        """
        Decorador para iniciar o SAP antes de executar a função decorada.

        :param f: Função a ser decorada.
        :return: Função decorada.
        """
        def wrap(self, *args, **kwargs):
            _self:SAPManipulation = self
            
            try:
                _self.session
            except AttributeError:
                _self.__conectar_sap()
            try:
                result =  f(_self, *args, **kwargs)
            finally:
                sleep(5)
                try:
                    if kwargs['fechar_sap_no_final']:
                        _self.fechar_sap()
                except:
                    pass
            return result
        return wrap
    
    # Decorador para verificar conexões
    @staticmethod
    def __verificar_conections(f):
        """
        Decorador para verificar conexões antes de executar a função decorada.

        :param f: Função a ser decorada.
        :return: Função decorada.
        """
        @wraps(f)
        def wrap(self, *args, **kwargs):
            _self:SAPManipulation = self
            
            result = f(_self, *args, **kwargs)
            try:
                if "Continuar com este logon sem encerrar os logons existentes".lower() in (choice:=_self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2")).text.lower():
                    choice.select()
                    _self.session.findById("wnd[0]").sendVKey(0)
            except:
                pass
            return result
        return wrap
        
    @__verificar_conections
    def __conectar_sap(self) -> None:
        """
        Conecta ao SAP, abrindo uma nova sessão se necessário.

        :raises Exception: Se não for possível conectar ao SAP.
        :raises ConnectionError: Se ocorrer um erro de conexão.
        """
        for _ in range(2):
            self.__session: win32com.client.CDispatch
            if not self.using_active_conection:
                try:
                    if not self.__verificar_sap_aberto():
                        subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                        sleep(5)
                    
                    SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
                    application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
                    
                    for _ in range(60*60):
                        try:
                            if self.__new_connection:
                                raise Exception("Erro controlado")
                            
                            conected_info = application.Children(0).Children(0).Info
                            if conected_info.SystemName.lower() != self.__ambiente.lower():# type: ignore
                                raise Exception("Erro controlado")
                            if conected_info.User.lower() != self.__user.lower():# type: ignore
                                raise Exception("Erro controlado")
                            
                            connection = application.Children(0) # type: ignore
                        except:
                            connection = application.OpenConnection(self.__ambiente, True) # type: ignore
                            self.__session = connection.Children(0)# type: ignore
                            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.__user # Usuario
                            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.__password # Senha
                            self.session.findById("wnd[0]").sendVKey(0)
                            break
                        
                            
                        if _ >= ((60*60) - 2):
                            Logs().register(status='Error', description="não foi possivel se conectar a mais uma tela do SAP", exception=traceback.format_exc())
                            sys.exit()
                        
                        if connection.Children.Count >= 6:
                            sleep(1)
                            continue
                        
                        novo_id = FindNewID(connection)
                        session = connection.Children(0)# type: ignore
                        
                        session.findById("wnd[0]").sendVKey(74)
                        
                        sleep(1)
                        self.__session = connection.Children(novo_id.target(connection))# type: ignore
                        break
                                    
                    try:
                        if (sbar:=self.session.findById("wnd[0]/sbar").text):
                            print(P(sbar, color="cyan"))
                    except:
                        pass
                    try:
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press() 
                    except:
                        pass
                    
                    return 
                
                
                except Exception as error:
                    if "sessão nao encontrada!" in str(error):
                        print(P("não foi possivel se conectar a mais uma tela do SAP", color='red'))
                        self.finalizar_programa_sap()
                        continue
                    if "connection = application.OpenConnection(self.__ambiente, True)" in traceback.format_exc():
                        raise Exception("SAP está fechado!")
                    else:
                        self.log.register(status='Error', description=str(error), exception=traceback.format_exc())
                        raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__class__} -> {error}")
            else:
                try:
                    if not self.__verificar_sap_aberto():
                        raise Exception("SAP está fechado!")
                    
                    self.SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
                    self.application: win32com.client.CDispatch = self.SapGuiAuto.GetScriptingEngine
                    self.connection: win32com.client.CDispatch = self.application.Children(0)
                    self.__session = self.connection.Children(0)
                    
                except Exception as error:
                    if "self.connection: win32com.client.CDispatch = self.application.Children(0)" in traceback.format_exc():
                        raise Exception("SAP está fechado!")
                    elif "SAP está fechado!" in traceback.format_exc():
                        raise Exception("SAP está fechado!")
                    else:
                        self.log.register(status='Error', description=str(error), exception=traceback.format_exc())

    # Método para fechar o SAP
    def fechar_sap(self):
        """
        Fecha a sessão atual do SAP.
        """
        print(P("fechando SAP!", color='red'))
        try:
            sleep(1)
            self.session.findById("wnd[0]").close()
            sleep(1)
            try:
                try:
                    self.session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
                except:
                    self.session.findById('wnd[2]/usr/btnSPOP-OPTION1').press()
            finally:
                del self.__session
        except Exception as error:
            print(P(f"não foi possivel fechar o SAP {type(error)} | {error}", color='red'))

    # Método para listar elementos
    @start_SAP
    def _listar(self, campo):
        """
        Lista os elementos de um campo específico.

        :param campo: Campo a ser listado.
        """
        cont = 0
        for child_object in self.session.findById(campo).Children:
            print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
            cont += 1

    # Método para verificar se o SAP está aberto
    def __verificar_sap_aberto(self) -> bool:
        """
        Verifica se o SAP está aberto.

        :return: True se o SAP estiver aberto, False caso contrário.
        """
        for process in psutil.process_iter(['name']):
            if "saplogon" in process.name().lower():
                return True
        return False    
    
    
    def finalizar_programa_sap(self):
        for proc in psutil.process_iter(['name']):
            if "sap" in proc.info['name'].lower():
                proc.kill()
                print("Processo SAP encerrado.")    
    
    # Método de teste         
    @start_SAP
    def _teste(self):
        """
        Método de teste para verificar a conexão com o SAP.
        """
        print("testado")
  
if __name__ == "__main__":
    pass
    #crd = Credential("SAP_QAS").load()
    
    #bot = SAPManipulation(user=crd['user'], password=crd['password'], ambiente="S4Q")
    #bot.conectar_sap()
    #bot._teste()
    
    #import pdb;pdb.set_trace()
    #bot.fechar_sap()
