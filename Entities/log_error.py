
import os
from getpass import getuser
from time import sleep
from datetime import datetime
import traceback

class LogError:
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
            with open(self.__path, "w")as _file:
                _file.write(f"data;tipo;descri;traceback\n")
    
    def register(self, tipo:type, descri:str, trace:str=" ") -> None:
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
                with open(self.__path, 'a')as file:
                    file.write(f"{datetime.now()};{tipo};{descri};{trace}\n")
                    return 
            except PermissionError:
                print(f"Feche o arquivo {self.__path} para que o registro seja feito")
                sleep(1)

if __name__ == "__main__":
    log_error = LogError(file="test.csv")
    try:
        raise Exception("teste")
    except Exception as error:
        log_error.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())

