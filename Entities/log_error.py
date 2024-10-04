
import os
from getpass import getuser
from time import sleep
from datetime import datetime
import traceback
import json

class LogError:
    informativo_path = ""
    
    def __init__(self, file:str="log_error.csv") -> None:
        '''
        construtur da classe define o caminho do arquivo e cria caso ele não exista
        
        Parameters:
        path: (str) = caminho do arquivo

        Return:
        None
        '''

        self.__path: str = f"C:\\Users\\{getuser()}\\.pagamentos_diarios\\"
        if not os.path.exists(self.__path):
            os.mkdir(self.__path)
        
        self.__path += file
        if not os.path.exists(self.__path):
            with open(self.__path, "w", encoding='utf-8')as _file:
                _file.write(f"data;tipo;descri;traceback\n")
    
    def register(self, tipo:type|str, descri:str, trace:str=" ") -> None:
        '''
        metodo para salvar o registro no arquivo .csv

        Parameters:
        tipo: (str) = tipo do erro 'geralmente estanciado como type(error)'
        descri: (str) = descrição do errir
        trace: (str) traceback col
        '''

        trace = trace.replace('\n', '|||')
        for x in range(5*60):
            try:
                with open(self.__path, 'a', encoding='utf-8')as file:
                    file.write(f"{datetime.now()};{tipo};{descri};{trace}\n")
                    return 
            except PermissionError:
                print(f"Feche o arquivo {self.__path} para que o registro seja feito")
                sleep(1)
    
    @staticmethod            
    def informativo(text:str):
        text = f"{datetime.now().strftime("[%d/%m/%Y - %H:%M:%S] ->  ")} {text}"
        print(text)
        if not LogError.informativo_path:
            return
        if not os.path.exists(LogError.informativo_path):
            with open(LogError.informativo_path, 'w', encoding='utf-8') as _file:
                json.dump([], _file)
        with open(LogError.informativo_path, 'r', encoding='utf-8') as _file:
            lista:list = json.load(_file)
        
        lista.append(text)
        
        with open(LogError.informativo_path, 'w', encoding='utf-8') as _file:
            json.dump(lista, _file)
        

if __name__ == "__main__":
    log_error = LogError(file="test.csv")
    try:
        raise Exception("teste")
    except Exception as error:
        log_error.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())

