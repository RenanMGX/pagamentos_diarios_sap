import configparser
import os
from typing import Literal, Dict
import sys

try:
    from Entities.dependencies.default_config import default as default_config
except ImportError:
    default_config: Dict[str,Dict[str,object]] = {}


class Config:
    @property
    def file_name(self) -> str:
        return "config.init"
    
    @property
    def config(self):
        return self.__config
    
    def __init__(self):
        if not os.path.exists(self.file_name):
            with open(self.file_name, 'w')as _file:
                _file.write("")
            self.__config = configparser.ConfigParser()
            self.read()
            if default_config:
                for key, options in default_config.items():
                    self.config.add_section(str(key))
                    if options:
                        for option, value in options.items():
                            self.config[str(key)][option] = str(value)
            self.__save()
            print(f"o arquivo '{self.file_name}' não existia então foi criado e o script sera encerrado!")
            sys.exit()
        else:
            self.__config = configparser.ConfigParser()
            self.read()
        
    def __getitem__(self, section:str):
        if self.config.has_section(str(section)):
            return self.config[str(section)]
        else:
            return {}
        
    def read(self):
        self.__config.read(self.file_name)
        
    def __save(self) -> None:
        with open(self.file_name, 'w')as _file:
            self.config.write(_file)
        self.read()
        
    def add(self, *, section:str, **kwargs):
        self.config.add_section(section)
        
        if kwargs:
            for key, value in kwargs.items():
                self.config[section][str(key)] = str(value)
            self.__save()
        else:
            raise Exception("nenhum atributo foi passado para alimentar o config")
    
    def alt(self, *, section:str, **kwargs):
        for key, value in kwargs.items():
            try:
                self.config[section][str(key)] = str(value)
            except Exception as error:
                print(type(error),str(error), f"---> {key=}:{value=}")
        self.__save()
        
    def delete(self, section:str, option:str="") -> None:
        if option:
            if self.config.has_option(section, option):
                self.config.remove_option(section, option)
            else:
                raise Exception(f"{option=} não foi encontrado")
        else:
            if self.config.has_section(section):
                self.config.remove_section(section)
            else:
                raise Exception(f"{section=} não foi encontrado")
        self.__save()
        
if __name__ == "__main__":
    config = Config()        
        