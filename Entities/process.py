class Processos:
    def __init__(self) -> None:
        self.__boleto:bool = False
        self.__consumo:bool = False
        self.__imposto:bool = False
        self.__darfs:bool = False
        self.__relacionais:bool = False
        
    def __str__(self) -> str:
       return f"{self.boleto=} \n{self.consumo=} \n{self.imposto=} \n{self.darfs=} \n{self.relacionais=}"

    def __repr__(self) -> str:
        return f"{self.boleto=} \n{self.consumo=} \n{self.imposto=} \n{self.darfs=} \n{self.relacionais=}"
    
    @property
    def boleto(self):
        return self.__boleto
    @boleto.setter
    def boleto(self, valor:bool):
        if not isinstance(valor, bool):
            raise TypeError("apenas valor booleano!")
        self.__boleto = valor
    
    @property
    def consumo(self):
        return self.__consumo
    @consumo.setter
    def consumo(self, valor:bool):
        if not isinstance(valor, bool):
            raise TypeError("apenas valor booleano!")
        self.__consumo = valor        
    
    @property
    def imposto(self):
        return self.__imposto
    @imposto.setter
    def imposto(self, valor:bool):
        if not isinstance(valor, bool):
            raise TypeError("apenas valor booleano!")
        self.__imposto = valor
        
    @property
    def darfs(self):
        return self.__darfs
    @darfs.setter
    def darfs(self, valor:bool):
        if not isinstance(valor, bool):
            raise TypeError("apenas valor booleano!")
        self.__darfs = valor
        
    @property
    def relacionais(self):
        return self.__relacionais
    @relacionais.setter
    def relacionais(self, valor:bool):
        if not isinstance(valor, bool):
            raise TypeError("apenas valor booleano!")
        self.__relacionais = valor

if __name__ == "__main__":
    pass