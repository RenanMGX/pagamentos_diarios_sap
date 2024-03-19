import json
import os
from copy import deepcopy
import traceback
from random import randint

class Credential:
    def __init__(self, path:str) -> None:
        exten = ".json"
        self.path = path+exten if not path.endswith(exten) else path
        
    def load(self) -> dict:
        """crie / ler um arquivo json contendo as credenciais

        Args:
            path (str): caminho do arquivo que será salvo as crendicias

        Raises:
            FileNotFoundError: caso o arquivo não exista irá criar um e vai alertar que foi criado e pedira para iniciar o script novamente

        Returns:
            dict: dicionario com a credenciais salvas
        """
        
        if not os.path.exists(self.path):
            with open(self.path, 'w')as _file:
                json.dump({"user": "", "password": "", "key": 0},_file)
            raise FileNotFoundError(f"{self.path=} não existe! então foi criar uma no repositorio, edite as credenciais e execute o codigo novamente!")

        with open(self.path, 'r')as _file:
            result:dict = json.load(_file)
        
        new_result = deepcopy(result)
        for key,value in new_result.items():
            if key == 'key':
                continue
            new_result[key] = self.decifrar(value, new_result['key'])
            
        return new_result
                            
    
    def save(self, *, user:str, password:str) -> None:
        key = randint(500,6000)
        with open(self.path, 'w')as _file:
            json.dump(
                {
                    'user':self.criar_cifra(user, key),
                    'password':self.criar_cifra(password, key),
                    'key': key
                },
                _file)
    
    def criar_cifra(self, text:str, key:int=1, response_json:bool=False) -> str:
        """criptografa a string informada orientada pela Key

        Args:
            text (str): texto a ser criptografado
            key (int, optional): chave para criptografia. Defaults to 1.
            response_json (bool, optional): retorna a string em formato json. Defaults to False.

        Returns:
            str: valor criptografado
        """
        if not isinstance(key, int):
            key = int(key)
        result:str = ""
        for letra in text:
            codigo:int = ord(letra) + key
            result += chr(codigo)
        
        if response_json:    
            return json.dumps(result)
        return result
    
    def decifrar(self, text:str, key:int) -> str:
        """descriptografa a string

        Args:
            text (str): texto a ser descriptografado
            key (int): chave para descriptografar

        Returns:
            str: texto descriptografado
        """
        return self.criar_cifra(text, -key)
        
if __name__ == "__main__":
    credential = Credential("creden/SAP")
    
    print(credential.load())