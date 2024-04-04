class Processos:
    def __init__(self) -> None:
        self.__boleto:bool = False
        self.__consumo:bool = False
        self.__imposto:bool = False
        
    def __str__(self) -> str:
       return f"{self.boleto=} \n{self.consumo=} \n{self.imposto=}"
    
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

if __name__ == "__main__":
    pass