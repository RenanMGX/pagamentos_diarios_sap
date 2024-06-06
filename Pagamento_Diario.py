from Entities.f110 import F110
from Entities.log_error import LogError
from Entities.process import Processos
#from Entities.rotinas import RotinasDB
from datetime import datetime
import traceback
import sys

from PyQt5 import QtCore, QtGui, QtWidgets

class Date:
    def __init__(self):
        self.date = 0


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.setEnabled(True)
        Dialog.resize(400, 400)

        font = QtGui.QFont()
        font.setKerning(True)

        Dialog.setFont(font)
        Dialog.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)#type: ignore
        Dialog.setAccessibleDescription("")
        Dialog.setLayoutDirection(QtCore.Qt.LeftToRight)#type: ignore
        Dialog.setInputMethodHints(QtCore.Qt.ImhNone)#type: ignore
        Dialog.setSizeGripEnabled(False)
        Dialog.setModal(False)

        self.w_calendario = QtWidgets.QCalendarWidget(Dialog)
        self.w_calendario.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))#type: ignore
        self.w_calendario.setGeometry(QtCore.QRect(40, 10, 312, 183))
        self.w_calendario.setInputMethodHints(QtCore.Qt.ImhNone)#type: ignore
        self.w_calendario.setGridVisible(False)
        self.w_calendario.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
        self.w_calendario.setNavigationBarVisible(True)
        self.w_calendario.setDateEditEnabled(True)
        self.w_calendario.setObjectName("w_calendario")
        
        hight_botton:int = 180
        width_botton:int = 40
        quant_checkbox:int = 0
        
        # self.label_quant_letras = QtWidgets.QLabel(Dialog)
        # self.label_quant_letras.setObjectName("boletos")
        # self.label_quant_letras.setGeometry(QtCore.QRect(290, 250, 200, 31))
        # self.label_quant_letras.setAlignment(QtCore.Qt.AlignCenter) # type: ignore
        # self.label_quant_letras.setText("Quantidade de letras Restantes:\n 17")
        
        
        self.cb_boletos = QtWidgets.QCheckBox(Dialog)
        self.cb_boletos.setObjectName("boletos")
        self.cb_boletos.setGeometry(QtCore.QRect(width_botton, (hight_botton + 25), 250, 31))
        self.cb_boletos.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor)) #type: ignore
        quant_checkbox += 1
        
        self.cb_consumo = QtWidgets.QCheckBox(Dialog)
        self.cb_consumo.setObjectName("boletos")
        self.cb_consumo.setGeometry(QtCore.QRect(width_botton, (hight_botton + 50), 250, 31))
        self.cb_consumo.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor)) #type: ignore
        quant_checkbox += 1
        
        self.cb_imposto = QtWidgets.QCheckBox(Dialog)
        self.cb_imposto.setObjectName("imposto")
        self.cb_imposto.setGeometry(QtCore.QRect(width_botton, (hight_botton + 75), 250, 31))
        self.cb_imposto.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor)) #type: ignore
        quant_checkbox += 1
        
        self.cb_darfs = QtWidgets.QCheckBox(Dialog)
        self.cb_darfs.setObjectName("DARFS, impostos Federais.")
        self.cb_darfs.setGeometry(QtCore.QRect(width_botton, (hight_botton + 100), 250, 31))
        self.cb_darfs.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor)) #type: ignore
        quant_checkbox += 1
        
        self.cb_relacionais = QtWidgets.QCheckBox(Dialog)
        self.cb_relacionais.setObjectName("Relacionais")
        self.cb_relacionais.setGeometry(QtCore.QRect(width_botton, (hight_botton + 125), 250, 31))
        self.cb_relacionais.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor)) #type: ignore
        quant_checkbox += 1
        
        self.bt_iniciar = QtWidgets.QPushButton(Dialog)
        self.bt_iniciar.setGeometry(QtCore.QRect(140, (hight_botton + (quant_checkbox * 25) + 40), 121, 31))
        self.bt_iniciar.setObjectName("bt_iniciar")
        self.bt_iniciar.clicked.connect(self.retornar_data)
        
        

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", f"Pagamentos Diarios - {version}"))
        self.bt_iniciar.setText(_translate("Dialog", "Iniciar"))
        
        self.cb_boletos.setText(_translate("Dialog", "Boletos - Consome 2 letras"))
        self.cb_consumo.setText(_translate("Dialog", "Consumo - Consome 1 letra"))
        self.cb_imposto.setText(_translate("Dialog", "Imposto - Consome 1 letra"))
        self.cb_darfs.setText(_translate("Dialog", "DARFS, impostos Federais - Consome 1 letra"))
        self.cb_relacionais.setText(_translate("Dialog", "Relacionais - Consome 1 letra"))

    def retornar_data(self):
        calendar_date = self.w_calendario.selectedDate()
        date.date = datetime(calendar_date.year(), calendar_date.month(), calendar_date.day())#type: ignore
        processo.boleto = self.cb_boletos.isChecked()
        processo.consumo = self.cb_consumo.isChecked()
        processo.imposto = self.cb_imposto.isChecked()
        processo.darfs = self.cb_darfs.isChecked()
        processo.relacionais = self.cb_relacionais.isChecked()
        Dialog.close()


if __name__ == "__main__":
    date = Date()
    processo = Processos()
    
    version = "v1.6"
    
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    #sys.exit(app.exec_())
    app.exec_()
    
    register_erro: LogError = LogError()
    finalizar:bool = False
    try:
        if date.date == 0:
            raise Exception("data invalida")
        bot: F110 = F110(date.date)#type: ignore
        bot.mostrar_datas()
        bot.iniciar(processo)
        print("\nScript Finalizado com exito!")
    except Exception as error:
        print("\nScript finalizado com o seguinte error")        
        print(f"{type(error)} -> {error}")
        register_erro.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())
    
        if str(error) == "data invalida":
            sys.exit()
            
    input("Digite algo para finalizar o Script: ")    
