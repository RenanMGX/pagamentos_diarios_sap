from datetime import datetime
from Entities.f110_auto import F110Auto
from Entities.log_error import LogError
from Preparar_Documentos_para_PGTO import Preparar
from Entities.crenciais import Credential
from Entities.process import Processos
import traceback
from getpass import getuser
import os
from dateutil.relativedelta import relativedelta

class PagamentosDiariosAuto(F110Auto):
    def __init__(self, *,user:str, password:str, ambiente:str, date:datetime) -> None:
        self.__date:datetime = date.replace(hour=0,minute=0,second=0,microsecond=0)
        
        F110Auto.__init__(self, date=self.date , user=user, password=password, ambiente=ambiente)
        
    @property
    def date(self):
        return self.__date


if __name__ == "__main__":
    try:
        param = {
            "qas" : ["S4Q","SAP_QAS"],
            "prd" :  ["S4P", "SAP_PRD"]
        }
        choose_param = "prd" #alterar entrada e ambiente SAP
        
        crd:dict = Credential(param[choose_param][1]).load()
        processos:Processos = Processos()
        processos.boleto = True
        processos.consumo = True
        processos.imposto = False    
        
        date:datetime = datetime.now()
        date = date.replace(hour=0,minute=0,second=0,microsecond=0)
        date = date + relativedelta(days=1)
        print(date)
        
        preparar = Preparar(date=date, arquivo_datas=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/Datas_Execução.xlsx")

        execute_program:bool = False
        for key,value in preparar.datas.items():
            if value['data_atual'] == date:
                execute_program = True
            
            if not execute_program:
                raise Exception("dia não permitido para execução do script")
        
        bot = PagamentosDiariosAuto(
            user=crd['user'],
            password=crd['password'],
            ambiente=param[choose_param][0],
            date=date,
        )
        
        #bot.mostrar_datas()
        bot.iniciar(processos, empresas_separada=["N013"], salvar_letra=True, fechar_sap_no_final=True)
    
    except:
        path:str = "logs/"
        if not os.path.exists(path):
            os.makedirs(path)
        file_name = path + f"LogError_{datetime.now().strftime('%d%m%Y%H%M%Y')}.txt"
        with open(file_name, 'w', encoding='utf-8')as _file:
            _file.write(traceback.format_exc())
