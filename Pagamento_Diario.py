from Entities.f110 import F110
from Entities.log_error import LogError
import traceback


if __name__ == "__main__":
    register_erro: LogError = LogError()
    try:
        bot: F110 = F110(int(input("dias: ")))
        bot.mostrar_datas()
        bot.iniciar()
    except Exception as error:
        print(f"{type(error)} -> {error}")
        register_erro.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())
    
    input()
