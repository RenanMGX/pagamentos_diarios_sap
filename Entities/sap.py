from getpass import getuser
import win32com.client
from datetime import datetime
import psutil
import subprocess
from time import sleep

class SAPManipulation():
    @property
    def ambiente(self):
        return self.__ambiente
    
    def __init__(self, *, user:str, password:str, ambiente:str) -> None:
        self.__user:str = user
        self.__password:str = password
        self.__ambiente:str = ambiente
        
        
    @property
    def session(self) -> win32com.client.CDispatch:
        return self.__session
    
    #decorador
    @staticmethod
    def usar_sap(f):
        def wrap(self, *args, **kwargs):
            try:
                self.session
            except AttributeError:
                self.conectar_sap()
            try:
                result =  f(self, *args, **kwargs)
            finally:
                sleep(5)
                try:
                    if kwargs['fechar_sap_no_final']:
                        self.fechar_sap()
                except:
                    pass
            return result
                #raise Exception("o sap precisa ser conectado primeiro!")
        return wrap
    
    @staticmethod
    def __verificar_conections(f):
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
    def conectar_sap(self) -> None:
        try:
            if not self._verificar_sap_aberto():
                subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                sleep(5)
            
            SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
            application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
            connection = application.OpenConnection(self.__ambiente, True) # type: ignore
            self.__session: win32com.client.CDispatch = connection.Children(0)# type: ignore
            
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.__user # Usuario
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.__password # Senha
            self.session.findById("wnd[0]").sendVKey(0)
            
            return 
        except Exception as error:
            raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__class__} -> {error}")

    #@usar_sap
    def fechar_sap(self):
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

    @usar_sap
    def _listar(self, campo):
        cont = 0
        for child_object in self.session.findById(campo).Children:
            print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
            cont += 1

    def _verificar_sap_aberto(self) -> bool:
        for process in psutil.process_iter(['name']):
            if "saplogon" in process.name().lower():
                return True
        return False    
    
    @usar_sap
    def _teste(self):
        print("testado")
    
    
if __name__ == "__main__":
    pass
    #crd = Credential("SAP_QAS").load()
    
    #bot = SAPManipulation(user=crd['user'], password=crd['password'], ambiente="S4Q")
    #bot.conectar_sap()
    #bot._teste()
    
    #import pdb;pdb.set_trace()
    #bot.fechar_sap()
