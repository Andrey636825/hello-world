import os
import sys

from PyQt5 import QtWidgets, uic, QtWebEngineWidgets
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QFont, QIcon, QPixmap
from PyQt5.QtCore import Qt, QModelIndex, pyqtSlot, QSize, QEvent, QThreadPool, pyqtSignal, QObject, QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView

import resources_rc
import subprocess
import webbrowser
import configparser

from PyRun import run, runHandle, excelHandle


class PdfWindow(QWebEngineView):
    def __init__(self, pdf_path: str):
        super().__init__()
        self.settings().setAttribute(
            QtWebEngineWidgets.QWebEngineSettings.PluginsEnabled, True)
        self.settings().setAttribute(
            QtWebEngineWidgets.QWebEngineSettings.PdfViewerEnabled, True)
        self.load(QUrl.fromUserInput('file:///help.pdf'))

        # s1 = 'file:///' + pdf_path.replace('\\', '/')
        # pass

        self.setGeometry(600, 50, 800, 600)


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('main_window.ui', self)

        self.treeview = self.findChild(QtWidgets.QTreeView, "mainTreeView")
        self.to_exel_push_button = self.findChild(QtWidgets.QPushButton, "toExcelPushButton")
        self.all_rec_check_box = self.findChild(QtWidgets.QCheckBox, "allRecordsCheckBox")
        self.run_push_button = self.findChild(QtWidgets.QPushButton, "runPushButton")
        self.info_push_button = self.findChild(QtWidgets.QPushButton, "infoPushButton")

        self.model = QStandardItemModel()
        self.index = QModelIndex()

        self.rh = runHandle.RunHandle()
        self.eh = excelHandle.ExcelHandle()

        self.exe_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
        self.last_dir = os.path.dirname(os.path.realpath(sys.argv[0]))

        config = configparser.ConfigParser()
        config.read('config.ini')
        self.acrobat_path = config['PDF_Path']['Acrobat_Path']

        self.to_excel_rows = []

        font = QFont()
        font.setWeight(QFont.Bold)
        self.item_run = QStandardItem(QIcon(QPixmap(":/icons/loading.png")), "Выполнение")
        self.item_run.setFont(font)
        self.item_to_excel = QStandardItem(QIcon(QPixmap(":/icons/excelSign.png")), "Вывод в Excel")
        self.item_to_excel.setFont(font)

        self.tree_values = [
            QStandardItem(QIcon(QPixmap(":/icons/automatic.png")), "Параметры"),
            QStandardItem(QIcon(QPixmap(":/icons/info.png")), "Инструкция пользователя"),
            QStandardItem(QIcon(QPixmap(":/icons/database.png")), "Операции с базой данных SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/database.png")), "Оптимизировать базу данных SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/database.png")), "Создать копию базы данных SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/database.png")), "Загрузить базу данных SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/database.png")), "Восстановить базу данных SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/repair.png")), "Тестирование и исправление XML файлов"),
            QStandardItem(QIcon(QPixmap(":/icons/xml.png")), "XML >> SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/xml.png")), "ФЭС"),
            QStandardItem(QIcon(QPixmap(":/icons/xml.png")), "LSOp  (3462-У)"),
            QStandardItem(QIcon(QPixmap(":/icons/xml.png")), "LSOZ  (3462-У)"),
            QStandardItem(QIcon(QPixmap(":/icons/xml.png")), "LSos  (3462-У)"),
            QStandardItem(QIcon(QPixmap(":/icons/123.png")), "Проверка уникальности номеров записей ФЭС"),
            QStandardItem(QIcon(QPixmap(":/icons/excelSign.png")), "Excel >> Excel"),
            QStandardItem(QIcon(QPixmap(":/icons/excelSign.png")), "Выписки по корр. счетам"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Excel >> SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "СПАРК (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "СПАРК (ИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Даты первых операций по счетам"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Даты регистраций ЮЛ"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Даты регистраций ИП"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Валютообменные операции"),
            QStandardItem(QIcon(QPixmap(":/icons/excel.png")), "Корр. счета"),
            QStandardItem(QIcon(QPixmap(":/icons/sql.png")), "SQL >> SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/sql.png")), "ИНН всех ЮЛ, ИП, ФЛ (ФЭС)"),
            QStandardItem(QIcon(QPixmap(":/icons/sql.png")), "ИНН клиентов банка (ФЭС)"),
            QStandardItem(QIcon(QPixmap(":/icons/sql.png")), "Дата регистрации клиентов банка (ФЭС)"),
            QStandardItem(QIcon(QPixmap(":/icons/sql.png")), "LSOp (3462-У). Клиенты банка"),
            QStandardItem(QIcon(QPixmap(":/icons/pdf.png")), "Pdf >> SQL"),
            QStandardItem(QIcon(QPixmap(":/icons/pdf.png")), "Настройка загрузки ЕГРЮЛ"),
            QStandardItem(QIcon(QPixmap(":/icons/pdf.png")), "Настройка загрузки ЕГРИП"),
            QStandardItem(QIcon(QPixmap(":/icons/pdf.png")), "ЕГРЮЛ / ЕГРИП"),
            QStandardItem(QIcon(QPixmap(":/icons/book.png")), "Словари"),
            QStandardItem(QIcon(QPixmap(":/icons/book.png")), "Редактирование словарей сравнения данных (ФЭС, СПАРК, ЕГРЮЛ, ЕГРИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/equal.png")), "Проверка идентичности данных"),
            QStandardItem(QIcon(QPixmap(":/icons/equal.png")), "Идентичность данных ЮЛ (сопоставление по ИНН)"),
            QStandardItem(QIcon(QPixmap(":/icons/equal.png")), "Идентичность данных ФЛ и ИП (сопоставление по ИНН)"),
            QStandardItem(QIcon(QPixmap(":/icons/equal.png")), "Идентичность данных ФЛ и ИП (сопоставление по ФИО)"),
            QStandardItem(QIcon(QPixmap(":/icons/equal.png")), "Идентичность данных ФЛ и ИП (сопоставление по ФИО и реквизитам документа)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Сравнение данных ФЭС с данными СПАРК, ЕГРЮЛ и ЕГРИП"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Незаполненные ИНН в ФЭС (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "КПП (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "ОГРН (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Дата регистрации (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Наименование (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Адрес (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "ИНН ЕИО (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "ФИО(Наименование) ЕИО (ЮЛ)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Незаполненные ИНН в ФЭС (ИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "ОГРНИП (ИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "Дата регистрации (ИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/lightning.png")), "ФИО (ИП)"),
            QStandardItem(QIcon(QPixmap(":/icons/one.png")), "Код 4006 (первая операция по счетам)"),
            QStandardItem(QIcon(QPixmap(":/icons/three.png")), "Код 4005 (три месяца с даты регистрации)"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Операции некоммерческих организаций"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Код 9001"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Код 9002"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Операции некоммерческих организаций (дополнительные коды)"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Код 9003"),
            QStandardItem(QIcon(QPixmap(":/icons/nko_.png")), "Код 9004"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Кассовые операции"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Код 1003 (валютообменные операции)"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Код 1004 (валютообменные операции)"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Кассовые операции (дополнительные коды)"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Код 1010 (получение)"),
            QStandardItem(QIcon(QPixmap(":/icons/ticket-office.png")), "Код 1011 (внесение)"),
            QStandardItem(QIcon(QPixmap(":/icons/money.png")), "Корр. счета (115-ФЗ ст.7.2)")
        ]
        self.parent_child = [
            (self.model, self.tree_values[0], "run"),  # "Параметры"
            (self.model, self.tree_values[1], None),  # "Инструкция пользователя"
            (self.model, self.tree_values[2], None),    # "Операции с базой данных SQL"
            (self.tree_values[2], self.tree_values[3], "run"),
            (self.tree_values[2], self.tree_values[4], "run"),
            (self.tree_values[2], self.tree_values[5], "run"),
            (self.tree_values[2], self.tree_values[6], "run"),
            (self.model, self.tree_values[7], "run + openFiles"),   # "Тестирование и исправление XML файлов"
            (self.model, self.tree_values[8], None),    # "XML >> SQL"
            (self.tree_values[8], self.tree_values[9], "run + toExcel + openFiles"),
            (self.tree_values[8], self.tree_values[10], "run + toExcel + openFiles"),
            (self.tree_values[8], self.tree_values[11], "run + toExcel + skip"),
            (self.tree_values[8], self.tree_values[12], "run + toExcel + skip"),
            (self.model, self.tree_values[13], "run"),   # "Проверка уникальности номера записи"
            (self.model, self.tree_values[14], None),  # "Excel >> Excel"
            (self.tree_values[14], self.tree_values[15], "run"),
            (self.model, self.tree_values[16], None),   # "Excel >> SQL"
            (self.tree_values[16], self.tree_values[17], "run"),
            (self.tree_values[16], self.tree_values[18], "run"),
            (self.tree_values[16], self.tree_values[19], "run"),
            (self.tree_values[16], self.tree_values[20], "run"),
            (self.tree_values[16], self.tree_values[21], "run"),
            (self.tree_values[16], self.tree_values[22], "run"),
            (self.tree_values[16], self.tree_values[23], "run"),
            (self.model, self.tree_values[24], "runThreads + toExcelSheets"),   # "SQL >> SQL"
            (self.tree_values[24], self.tree_values[25], "run + toExcel"),
            (self.tree_values[24], self.tree_values[26], "run + toExcel"),
            (self.tree_values[24], self.tree_values[27], "run + toExcel"),
            (self.tree_values[24], self.tree_values[28], "run + toExcel"),
            (self.model, self.tree_values[29], None),   # "Pdf >> SQL"
            (self.tree_values[29], self.tree_values[30], "run + dialog"),
            (self.tree_values[29], self.tree_values[31], "run + dialog"),
            (self.tree_values[29], self.tree_values[32], "run + toExcel"),
            (self.model, self.tree_values[33], None),   # "Словари"
            (self.tree_values[33], self.tree_values[34], "run + dialog"),
            (self.model, self.tree_values[35], "runThreads + toExcelSheets"),   # "Проверка идентичности данных"
            (self.tree_values[35], self.tree_values[36], "run + toExcel"),
            (self.tree_values[35], self.tree_values[37], "run + toExcel"),
            (self.tree_values[35], self.tree_values[38], "run + toExcel"),
            (self.tree_values[35], self.tree_values[39], "run + toExcel"),
            (self.model, self.tree_values[40], "runThreads + toExcelSheets"),   # "Сравнение данных ФЭС с данными СПАРК, ЕГРЮЛ и ЕГРИП"
            (self.tree_values[40], self.tree_values[41], "run + toExcel"),
            (self.tree_values[40], self.tree_values[42], "run + toExcel"),
            (self.tree_values[40], self.tree_values[43], "run + toExcel"),
            (self.tree_values[40], self.tree_values[44], "run + toExcel"),
            (self.tree_values[40], self.tree_values[45], "run + toExcel"),
            (self.tree_values[40], self.tree_values[46], "run + toExcel"),
            (self.tree_values[40], self.tree_values[47], "run + toExcel"),
            (self.tree_values[40], self.tree_values[48], "run + toExcel"),
            (self.tree_values[40], self.tree_values[49], "run + toExcel"),
            (self.tree_values[40], self.tree_values[50], "run + toExcel"),
            (self.tree_values[40], self.tree_values[51], "run + toExcel"),
            (self.tree_values[40], self.tree_values[52], "run + toExcel"),
            (self.model, self.tree_values[53], "run + toExcel"),    # "Код 4006 (первая операция по счетам)"
            (self.model, self.tree_values[54], "run + toExcel"),    # "Код 4005 (три месяца с даты регистрации)"
            (self.model, self.tree_values[55], "runThreads + toExcelSheets"),   # "Операции некоммерческих организаций"
            (self.tree_values[55], self.tree_values[56], "run + toExcel"),
            (self.tree_values[55], self.tree_values[57], "run + toExcel"),
            (self.model, self.tree_values[58], "runThreads + toExcelSheets"),   # "Операции некоммерческих организаций (дополнительные коды)"
            (self.tree_values[58], self.tree_values[59], "run + toExcel"),
            (self.tree_values[58], self.tree_values[60], "run + toExcel"),
            (self.model, self.tree_values[61], "runThreads + toExcelSheets"),   # "Кассовые операции"
            (self.tree_values[61], self.tree_values[62], "run + toExcel"),
            (self.tree_values[61], self.tree_values[63], "run + toExcel"),
            (self.model, self.tree_values[64], "runThreads + toExcelSheets"),   # "Кассовые операции (дополнительные коды)"
            (self.tree_values[64], self.tree_values[65], "run + toExcel"),
            (self.tree_values[64], self.tree_values[66], "run + toExcel"),
            (self.model, self.tree_values[67], "run + toExcel")     # "Корр. счета (115-ФЗ ст.7.2)"
        ]
        for i in self.parent_child:
            if (i[2] is not None) and ("skip" in i[2]):
                continue
            i[0].appendRow(i[1])

        self.treeview.setHeaderHidden(True)
        self.treeview.setFont(QFont('MS Shell Dig 2', 10))
        self.treeview.setModel(self.model)
        self.treeview.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.treeview.clicked[QModelIndex].connect(self.on_treeview_clicked)
        self.treeview.setIconSize(QSize(25, 25))

        self.run_push_button.clicked.connect(self.on_run_button_clicked)
        self.to_exel_push_button.clicked.connect(self.on_to_excel_button_clicked)
        self.info_push_button.clicked.connect(self.on_info_button_clicked)

        self.treeview.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.ContextMenu:
            exp_action = QtWidgets.QAction(QIcon(QPixmap(":/icons/expandAll.png")), "Показать", self)
            exp_action.triggered.connect(lambda: self.treeview.expandAll())
            col_action = QtWidgets.QAction(QIcon(QPixmap(":/icons/collapseAll.png")), 'Скрыть', self)
            col_action.triggered.connect(lambda: self.treeview.collapseAll())

            menu = QtWidgets.QMenu(self)
            menu.addAction(exp_action)
            menu.addAction(col_action)
            menu.setFont(QFont('MS Shell Dig 2', 10))
            menu.exec_(event.globalPos())
            return True
        return False

    @pyqtSlot(QModelIndex)
    def on_treeview_clicked(self, index):
        if not index.isValid():
            self.info_push_button.setEnabled(False)
            self.to_exel_push_button.setEnabled(False)
            self.all_rec_check_box.setEnabled(False)
            self.run_push_button.setEnabled(False)
            self.to_exel_push_button.setText("  Вывести в Excel  ")
            self.run_push_button.setText("Выполнить  ")
            return

        self.index = index
        item = self.model.itemFromIndex(self.index)
        p = self.tree_values.index(item)

        print(item.text())

        self.info_push_button.setEnabled(True)

        self.to_exel_push_button.setEnabled(False)
        self.all_rec_check_box.setEnabled(False)
        self.run_push_button.setEnabled(False)
        if self.parent_child[p][2] is None:
            self.to_exel_push_button.setText("  Вывести в Excel  ")
            self.run_push_button.setText("Выполнить  ")
        else:
            if "runThread" in self.parent_child[p][2]:
                self.run_push_button.setEnabled(True)
                self.run_push_button.setText("Выполнить  . . .  ")
            elif "run" in self.parent_child[p][2]:
                self.run_push_button.setEnabled(True)
                self.run_push_button.setText("Выполнить  ")
            if "toExcelSheets" in self.parent_child[p][2]:
                self.to_exel_push_button.setEnabled(True)
                self.to_exel_push_button.setText("  Вывести в Excel  . . .  ")
                self.all_rec_check_box.setEnabled(True)
            elif "toExcel" in self.parent_child[p][2]:
                self.to_exel_push_button.setEnabled(True)
                self.to_exel_push_button.setText("  Вывести в Excel  ")
                self.all_rec_check_box.setEnabled(True)

    @pyqtSlot()
    def on_run_button_clicked(self):
        if not self.index.isValid():
            return

        if QtWidgets.QMessageBox.question(self,"", "Выполнить выбранный пункт?",
                                          QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.No) == QtWidgets.QMessageBox.No:
            return

        item = self.model.itemFromIndex(self.index)
        p = self.tree_values.index(item)

        file_name = ""
        file_names = []
        if "openFiles" in self.parent_child[p][2]:
            options = QtWidgets.QFileDialog.Options()
            options |= QtWidgets.QFileDialog.DontUseNativeDialog
            file_names, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Выбрать  xml-файл", self.last_dir,
                                                                 "xml-файлы  (*.xml);;Все  файлы  (*)", options=options)
            if len(file_names) < 1:
                return
            else:
                self.last_dir = os.path.dirname(os.path.abspath(file_names[0]))

        self.rh.set_file_names(file_names)

        if not ("dialog" in self.parent_child[p][2]):
            run_dialog = run.RunDialog(self)
            self.rh.set_output_model(run_dialog.list_view.model())
            run_dialog.write_item(self.item_run.clone())
            run_dialog.write_item(item.clone())
            #
            for f in self.rh.run_func_list[p]:
                runnable = run.Run(f)
                QThreadPool.globalInstance().start(runnable)
            #
            run_dialog.start_timer()
            run_dialog.exec()
        else:
            for f in self.rh.run_func_list[p]:
                runnable = run.Run(f)
                QThreadPool.globalInstance().start(runnable)

    @pyqtSlot()
    def on_to_excel_button_clicked(self):
        if not self.index.isValid():
            return

        if QtWidgets.QMessageBox.question(self,"", "Вывести информацию в Excel?",
                                          QtWidgets.QMessageBox.Yes, QtWidgets.QMessageBox.No) == QtWidgets.QMessageBox.No:
            return

        if not self.all_rec_check_box.isChecked():
            to_excel_dialog = run.ToExcelDialog(self, self.to_excel_rows)
            to_excel_dialog.setWindowIcon(QIcon(QPixmap(":/icons/excelSign.png")))
            to_excel_dialog.exec()
            if not to_excel_dialog.ok_pressed:
                return
            self.eh.set_output_rows(self.to_excel_rows)
        else:
            self.eh.set_output_rows([])

        item = self.model.itemFromIndex(self.index)
        p = self.tree_values.index(item)

        run_dialog = run.RunDialog(self)
        self.eh.set_output_model(run_dialog.list_view.model())
        run_dialog.write_item(self.item_to_excel.clone())
        run_dialog.write_item(item.clone())
        #
        for f in self.eh.excel_func_list[p]:
            runnable = run.Run(f)
            QThreadPool.globalInstance().start(runnable)
        #
        run_dialog.start_timer()
        run_dialog.exec()

    def on_info_button_clicked(self):
        process = subprocess.Popen(['C:/Program Files/Adobe/Acrobat DC/Acrobat/Acrobat.exe', '/A',
                                   'page=3', self.exe_dir + '/docs/help.pdf'])
        # process.wait()

        # url = self.exe_dir + '/docs/help.htm'
        # webbrowser.open(url)

        # w = PdfWindow(self.exe_dir + '/docs/help.pdf')
        # w.show()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
