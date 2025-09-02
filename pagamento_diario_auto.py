from datetime import datetime
from Entities.f110_auto import F110Auto
from Entities.log_error import LogError
from Entities.dependencies.logs import Logs
from Preparar_Documentos_para_PGTO import Preparar
from Entities.crenciais import Credential
from Entities.process import Processos
import traceback
from getpass import getuser
import os
from dateutil.relativedelta import relativedelta
from time import sleep
from typing import List,Literal,Dict
import pandas as pd
import json
import sys

class PagamentosDiariosAuto(F110Auto):
    def __init__(self, *,user:str, password:str, ambiente:str, date:datetime) -> None:
        self.__date:datetime = date.replace(hour=0,minute=0,second=0,microsecond=0)
        
        F110Auto.__init__(self, date=self.date , user=user, password=password, ambiente=ambiente)
        
    @property
    def date(self):
        return self.__date

if __name__ == "__main__":
    LogError.informativo_path = os.path.join(os.getcwd(), 'informativo_pgmt_diario.json')
    
    if os.path.exists(LogError.informativo_path):
        os.unlink(LogError.informativo_path)
    for _ in range(1):
        try:
            param:dict = {
                "qas" : ["S4Q","SAP_QAS"],
                "prd" :  ["S4P", "SAP_PRD"],
                "django": ["S4P", "SAP_PRD"],
            }
            
            processos:Processos = Processos()
            empresas:list = [] # LIMPAR ESSA LISTA PARA PRODUÇÂO
                    
            choose_param:Literal["qas", "prd", "django"] = 'prd' #alterar entrada e ambiente SAP
            
            django_argv_path = 'django_argv.json'
            if os.path.exists(django_argv_path):
                try:
                    LogError.informativo("iniciando via Django <django:blue>")
                    
                    with open(django_argv_path, 'r', encoding='utf-8')as _file:
                        argvs:dict = json.load(_file)
                    
                    os.unlink(django_argv_path)    
                    
                    if argvs.get('date'):
                        date = datetime.strptime(str(argvs.get('date')), '%Y-%m-%d')
                    else:
                        raise Exception("data não encontrada")
                        
                    if 'boleto' in argvs:
                        processos.boleto = bool(argvs.get('boleto'))
                    else:
                        processos.boleto = True
                        
                    if 'consumo' in argvs:
                        processos.consumo = bool(argvs.get('consumo'))
                    else:
                        processos.consumo = True
                        
                    if 'imposto' in argvs:
                        processos.imposto = bool(argvs.get('imposto'))
                    else:
                        processos.imposto = True
                        
                    if 'darfs' in argvs:
                        processos.darfs = bool(argvs.get('darfs'))
                    else:
                        processos.darfs = True
                        
                    if 'relacionais' in argvs:
                        processos.relacionais = bool(argvs.get('relacionais'))
                    else:
                        processos.relacionais = True
                        
                    if argvs.get('empresas'):
                        empresas = argvs['empresas']
                        print(empresas)
                    
                    choose_param = 'django'
                except Exception as err:
                    LogError.informativo(f"{str(traceback.format_exc())} <django:red>")
                    sys.exit()
                
            else:
                date:datetime = datetime.now()
                date = date.replace(hour=0,minute=0,second=0,microsecond=0)
                date = (date + relativedelta(days=0)) if choose_param == "prd" else (date - relativedelta(days=0))

            print(f"{'#'*100}\nExecutando em TESTES\n{'#'*100}") if choose_param == "qas" else print(f"{'#'*100}\nExecutando em PRODUÇÃO\n{'#'*100}") if choose_param == "prd" else print(f"{'#'*100}\nEXECUTÇÃO NÃO IDENTIFICADA - {choose_param}\n{'#'*100}")
                            
            crd:dict = Credential(param[choose_param][1]).load()

            preparar = Preparar(date=date, arquivo_datas=f"C:/Users/{getuser()}/PATRIMAR ENGENHARIA S A/RPA - Documentos/RPA - Dados/Pagamentos Diarios - Contas a Pagar/Datas_Execução.xlsx")

            execute_program:bool = False
            for key,value in preparar.datas.items():
                if value['data_atual'] == date:
                    execute_program = True
                
                if not execute_program:
                    raise Exception(f"a data selecionada {date} não é permitida para execução do script")
            
            bot = PagamentosDiariosAuto(
                user=crd['user'],
                password=crd['password'],
                ambiente=param[choose_param][0],
                date=date,
            )

            if choose_param == "qas":
                processos.boleto = True
                processos.consumo = False
                processos.imposto = False 
                processos.darfs = False
                processos.relacionais = True  
                
                bot.iniciar(processos,  salvar_letra=True, fechar_sap_no_final=True, empresas_separada=["P018"])
            
            elif choose_param == 'django':
                if empresas:
                    bot.iniciar(processos, salvar_letra=True, fechar_sap_no_final=True , empresas_separada=empresas)
                else:
                    bot.iniciar(processos, salvar_letra=True, fechar_sap_no_final=True)# , empresas_separada=["N017"])
                
            else: # ==== "prd"
                processos.boleto = True
                processos.consumo = True
                processos.imposto = True 
                processos.darfs = True
                processos.relacionais = True
                
                if empresas:
                    bot.iniciar(processos, salvar_letra=True, fechar_sap_no_final=True , empresas_separada=empresas)
                else:                    
                    bot.iniciar(processos, salvar_letra=True, fechar_sap_no_final=True)# , empresas_separada=["N017"])
            
            try:
                os.unlink(bot.nome_arquivo)
            except:
                pass
            
            LogError.informativo("Automação Finalizada com Sucesso! <django:green>")
            Logs().register(status='Concluido', description="Automação Finalizada com Sucesso!")
        
        except Exception as error:
            Logs().register(status='Error', description=str(error), exception=traceback.format_exc())
            
            LogError.informativo(f"{str(error)} <django:red>")
            path:str = "logs/"
            if not os.path.exists(path):
                os.makedirs(path)
            file_name = path + f"LogError_{datetime.now().strftime('%d%m%Y%H%M%Y')}.txt"
            with open(file_name, 'w', encoding='utf-8')as _file:
                _file.write(traceback.format_exc())
            raise error
        sleep(1)
