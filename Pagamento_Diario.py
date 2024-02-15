from Entities.f110 import F110
from Entities.log_error import LogError
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
        Dialog.resize(400, 300)

        font = QtGui.QFont()
        font.setKerning(True)

        Dialog.setFont(font)
        Dialog.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        Dialog.setAccessibleDescription("")
        Dialog.setLayoutDirection(QtCore.Qt.LeftToRight)
        Dialog.setInputMethodHints(QtCore.Qt.ImhNone)
        Dialog.setSizeGripEnabled(False)
        Dialog.setModal(False)

        self.w_calendario = QtWidgets.QCalendarWidget(Dialog)
        self.w_calendario.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.w_calendario.setGeometry(QtCore.QRect(40, 10, 312, 183))
        self.w_calendario.setInputMethodHints(QtCore.Qt.ImhNone)
        self.w_calendario.setGridVisible(False)
        self.w_calendario.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
        self.w_calendario.setNavigationBarVisible(True)
        self.w_calendario.setDateEditEnabled(True)
        self.w_calendario.setObjectName("w_calendario")

        self.bt_iniciar = QtWidgets.QPushButton(Dialog)
        self.bt_iniciar.setGeometry(QtCore.QRect(130, 240, 121, 31))
        self.bt_iniciar.setObjectName("bt_iniciar")
        self.bt_iniciar.clicked.connect(self.retornar_data)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.bt_iniciar.setText(_translate("Dialog", "Iniciar"))

    def retornar_data(self):
        calendar_date = self.w_calendario.selectedDate()
        date.date = datetime(calendar_date.year(), calendar_date.month(), calendar_date.day())
        Dialog.close()


if __name__ == "__main__":

    date = Date()

    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    #sys.exit(app.exec_())
    app.exec_()


    register_erro: LogError = LogError()
    try:
        if date.date == 0:
            raise Exception("data invalida")
        bot: F110 = F110(date.date)
        bot.mostrar_datas()
        bot.iniciar()
        print("\nScript Finalizado com exito!")
    except Exception as error:
        print("\nScript finalizado com o seguinte error")        
        print(f"{type(error)} -> {error}")
        register_erro.register(tipo=type(error), descri=str(error), trace=traceback.format_exc())
    input()
    
