
import os
import json
import shutil
from datetime import datetime
from getpass import getuser
import pandas as pd
import mysql.connector as mysql
from dateutil.relativedelta import relativedelta
from copy import deepcopy

try:
    from Entities.db_credencial import crd as db_crd
except:
    from db_credencial import crd as db_crd

def verificarData(data: datetime, caminho: str) -> bool:
        data = data.replace(hour=0,minute=0,second=0,microsecond=0)
        df = pd.read_excel(caminho)[['Data', 'Execução']]
        dicio = {x[0]: x[1] for x in df.values}
        if dicio[data].lower() == 'sim':
            return True
        return False

class Rotinas:
    def __init__(self, data) -> None:
        self.__data: datetime = data
        self.__arquivo: str = "controle_pagamento_diario.json"
        self.__caminho_servidor: str = "\\\\server008\\G\\ARQ_PATRIMAR\\WORK\\TI - RPA\\#controle_scripts\\pagamentos_diarios\\"
        self.__caminho_local: str = f"C:\\Users\\{getuser()}\\.pagamentos_diarios\\"

        self.__possiveis_rotinas = [
            ["w", "y"],
            ["v", "x"],
            ["s", "u"],
            ["r", "t"],
            ["o", "q"],
            ["n", "p"],
            ["k", "m"],
            ["j", "l"],
            ["g", "i"],
            ["f", "h"]
        ]

        #criar a pasta no local
        if not os.path.exists(self.__caminho_local):
            os.makedirs(self.__caminho_local)
        
        #online
        if os.path.exists(self.__caminho_servidor):
            if not os.path.exists(self.__caminho_servidor + self.__arquivo):
                with open(self.__caminho_servidor + self.__arquivo, 'w')as file:
                    json.dump([], file)
                    shutil.copy2(self.__caminho_servidor + self.__arquivo, self.__caminho_local + self.__arquivo)
        #offline
        else:
            if not os.path.exists(self.__caminho_local + self.__arquivo):
                with open(self.__caminho_local + self.__arquivo, 'w')as file:
                    json.dump([], file)

    def ler(self) -> dict:
        if os.path.exists(self.__caminho_servidor + self.__arquivo):
            with open(self.__caminho_servidor + self.__arquivo, 'r')as _file:
                print("carregando arquivo Online:")
                retorno = [x for x in json.load(_file) if x['data'] == self.__data.strftime("%d/%m/%Y")]
                return retorno[0]
        else:
            if os.path.exists(self.__caminho_local + self.__arquivo): 
                print("carregando arquivo Offline:")
                retorno = [x for x in json.load(_file) if x['data'] == self.__data.strftime("%d/%m/%Y")] #type: ignore
                return retorno[0]
        
        return {}
    
    def proxima_rotina(self) -> list:
        data = self.__data.strftime("%d/%m/%Y")
        print(data)
        rotinas: list = []
        #online
        if os.path.exists(self.__caminho_servidor):
            with open(self.__caminho_servidor + self.__arquivo, 'r')as file:
                rotinas = json.load(file)

            rotinas_hj = [x for x in rotinas if x['data'] == data]
            if not rotinas_hj:
                rotinas.append({'data' : data, "rotina": [self.__possiveis_rotinas[0]]})
                with open(self.__caminho_servidor + self.__arquivo, 'w')as file:
                    json.dump(rotinas, file)
                shutil.copy2(self.__caminho_servidor + self.__arquivo, self.__caminho_local + self.__arquivo)
                return self.__possiveis_rotinas[0]

            quantidade_rotinas_diaria = len(rotinas_hj[0]['rotina'])

            try:
                rotinas[0]['rotina'].append(self.__possiveis_rotinas[quantidade_rotinas_diaria])
            except IndexError:
                raise Exception(f"quantidade maxima de rotinas execida neste dia '{data}'")


            with open(self.__caminho_servidor + self.__arquivo, 'w')as file:
                json.dump(rotinas, file)
            shutil.copy2(self.__caminho_servidor + self.__arquivo, self.__caminho_local + self.__arquivo)

            return self.__possiveis_rotinas[quantidade_rotinas_diaria]

        #offline
        else:
            with open(self.__caminho_local + self.__arquivo, 'r')as file:
                rotinas = json.load(file)

            rotinas_hj = [x for x in rotinas if x['data'] == data]
            if not rotinas_hj:
                rotinas.append({'data' : data, "rotina": [self.__possiveis_rotinas[0]]})
                with open(self.__caminho_local + self.__arquivo, 'w')as file:
                    json.dump(rotinas, file)
                return self.__possiveis_rotinas[0]

            quantidade_rotinas_diaria = len(rotinas_hj[0]['rotina'])

            try:
                rotinas[0]['rotina'].append(self.__possiveis_rotinas[quantidade_rotinas_diaria])
            except IndexError:
                raise Exception(f"quantidade maxima de rotinas execida neste dia '{data}'")

            rotinas[0]['rotina'].append(self.__possiveis_rotinas[quantidade_rotinas_diaria])
            with open(self.__caminho_local + self.__arquivo, 'w')as file:
                json.dump(rotinas, file)

            return self.__possiveis_rotinas[quantidade_rotinas_diaria]

class RotinasDB:
    def __init__(self, date:datetime=datetime.now()) -> None:
        self.__date:datetime = date
        self.__crd:dict = db_crd
        self.__rotinas_letras:list = [chr(101 + num) for num in range(22)]
    
    @property
    def date(self):
        return self.__date
    
    @property
    def crd(self):
        return self.__crd
    
    @property
    def rotinas_letras(self):
        return self.__rotinas_letras
        
    def load(self) -> list:
        connection = mysql.connect(
            host=self.crd['host'],
            user=self.crd['user'],
            password=self.crd['password'],
            database=self.crd['database']
        )
        cursor = connection.cursor()
        cursor.execute(f"SELECT rotina FROM rotinas WHERE date='{self.date.strftime('%Y-%m-%d')}'")
        
        letras:list = [letra[0].lower() for letra in cursor.fetchall()]#type: ignore
                
        connection.close()
        
        return letras
    
    def available(self, use_and_save:bool=False, all:bool=False) -> str:
        letras_disponiveis = deepcopy(self.rotinas_letras)
        for letra in self.load():
            try:
                letras_disponiveis.pop(letras_disponiveis.index(letra))
            except:
                continue
        
        if len(letras_disponiveis) > 0:
            if use_and_save:
                self.save_utilized(letter=letras_disponiveis[-1])  
            if not all:
                return letras_disponiveis[-1]
            else:
                return str(letras_disponiveis)
        raise Exception("sem letras disponiveis")
    
    def save_utilized(self, *, letter) -> None:
        connection = mysql.connect(
            host=self.crd['host'],
            user=self.crd['user'],
            password=self.crd['password'],
            database=self.crd['database']
        )
        cursor = connection.cursor()
        cursor.execute(f"INSERT INTO rotinas(date, rotina) VALUES ('{self.date.strftime('%Y-%m-%d')}', '{letter}')")
        connection.commit()
        connection.close()

if __name__ == "__main__":
    # procurar_rotinas = Rotinas(datetime.now())

    # print(procurar_rotinas.ler())
    # print(verificarData(data=datetime.now(), caminho=".TEMP/Datas_Execução.xlsx"))
    bot = RotinasDB(date=datetime.now()-relativedelta(days=0))

    letr = bot.available(all=True)
    print(letr)
    #bot.save_utilized(letter=letr)
    