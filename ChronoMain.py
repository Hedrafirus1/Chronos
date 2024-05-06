from PyQt5.QtWidgets import (
    QMainWindow,
    QApplication,
    QLabel,
    QPushButton,
    QFrame,
    QLineEdit,
    QLCDNumber,
    QTextEdit,
    QProgressBar,
    QSpinBox,
    QMenuBar,
    QStatusBar,
    QFileDialog,
)
import openpyxl as opxl
import random
from PyQt5 import uic
import numpy as np
import sys
import time


class UI(QMainWindow):
    def __init__(self):

        super(UI, self).__init__()
        self.core_Object = Core(self)
        self.timer = time.time()
        # Load ui file
        uic.loadUi("Chronos_interface.ui", self)

        # dictionary_cypher_and_proffesion(key:value)
        self.alternative_codes = {
            11629: [11735, 15511, 16211],
            11735: [11629],
            12460: [12582],
            12518: [12582],
            12612: [12582],
            12837: [12912, 12582, 16045],
            12855: [12912, 12582, 16045],
            12950: [12912, 12582, 16045],
            12991: [12912, 12582, 16045],
            13063: [17405, 17787],
            13395: [13049, 13419, 13450],
            13399: [18562],
            13419: [14618],
            13450: [18562],
            13462: [15452],
            15023: [15027],
            15287: [15236, 15452],
            15511: [15916, 12582],
            15914: [12582],
            15916: [12582],
            16464: [14618],
            17008: [18466, 18452],
            17209: [17405],
            17928: [17958, 17986, 18466],
            18338: [18336],
            18933: [18569],
            19100: [18569],
            19163: [19149],
            19182: [19100],
            19293: [17405, 18569],
            19700: [18466, 18452],
            19973: [12582],
        }
        self.proffesions = {
            11629: "Гальваник",
            11735: "Гравер",
            12460: "Изготовитель трафаретов, шкал и плат",
            12518: "Измеритель ЭФП",
            12582: "Испытатель деталей и приборов",
            12612: "Испытатель электромашин, апп. и приборов",
            12837: "Комплектовщик",
            12855: "Комплектовщик изделий",
            12950: "Контролёр деталей и приборов",
            12991: "Контролёр материалов, металлов, полуфабрикатов и изделий",
            13053: "Контролёр сборки электромашин аппаратов и приборов",
            13063: "Контролёр станочных и слесарных работ",
            13395: "Литейщик на машинах для литья под давлением",
            13399: "Литейщик пластмасс",
            13419: "Лудильщик деталей и приборов Г/С",
            13450: "Маляр",
            13462: "Маркировщик деталей и приборов",
            14618: "Монтажник РЭАиП",
            15023: "Намотчик катушек",
            15287: "Обработчик изделий из пластмасс",
            15452: "Окрасчик приборов и деталей",
            15511: "Оператор вакуумно-напылительных процессов",
            15914: "Оператор прецизионной резки",
            15916: "Оператор прецизионной фотолитографии",
            16045: "Оператор станков с ПУ",
            16464: "Паяльщик радиодеталей",
            17008: "Прессовщик изделий из пластмасс",
            17209: "Приготовитель растворов и смесей",
            17405: "Промывщик деталей и узлов",
            17861: "Регулировщик РЭАиП",
            17928: "Резчик на пилах, ножовках и станках",
            18193: "Сборщик микросхем",
            18312: "Сборщик электрических машин и аппаратов",
            18336: "Сварщик на лазерных установках",
            18338: "Сварщик на машинах контактной сварки",
            18466: "Слесарь механосборочных работ",
            18569: "Слесарь-сборщик РЭАиП",
            18874: "Столяр",
            18933: "Сушильщик детали и приборов",
            19100: "Термист",
            19149: "Токарь",
            19163: "Токарь-расточник",
            19182: "Травильщик",
            19293: "Укладчик-упаковщик",
            19479: "Фрезеровщик",
            19630: "Шлифовщик",
            19700: "Штамповщик",
            19940: "Электроэрозтонист",
            19973: "Юстировщик деталей и приборов",
            24013: "Мастер цеха",
        }

        # Define widgets

        self.lineEdit_Main = self.findChild(QLineEdit, "lineEdit_Main")
        self.label_5 = self.findChild(QLabel, "label_5")
        self.pushButton_Main = self.findChild(QPushButton, "pushButton_Main")
        self.frame_2 = self.findChild(QFrame, "frame_2")
        self.lineEdit_Marsh = self.findChild(QLineEdit, "lineEdit_Marsh")
        self.pushButton_Marsh = self.findChild(QPushButton, "pushButton_Marsh")
        self.label = self.findChild(QLabel, "label")
        self.frame_4 = self.findChild(QFrame, "frame_4")
        self.lineEdit_Prof = self.findChild(QLineEdit, "lineEdit_Prof")
        self.label_2 = self.findChild(QLabel, "label_2")
        self.pushButton_Prof = self.findChild(QPushButton, "pushButton_Prof")
        self.frame_5 = self.findChild(QFrame, "frame_5")
        self.lineEdit_Rand = self.findChild(QLineEdit, "lineEdit_Rand")
        self.pushButton_Rand = self.findChild(QPushButton, "pushButton_Rand")
        self.frame_6 = self.findChild(QFrame, "frame_6")
        self.lineEdit_Save = self.findChild(QLineEdit, "lineEdit_Save")
        self.pushButton_Save = self.findChild(QPushButton, "pushButton_Save")
        self.label_4 = self.findChild(QLabel, "label_4")
        self.frame = self.findChild(QFrame, "frame")
        self.lineEdit_Watcher2 = self.findChild(QLineEdit, "lineEdit_Watcher2")
        self.label_14 = self.findChild(QLabel, "label_14")
        self.lineEdit_Watcher1 = self.findChild(QLineEdit, "lineEdit_Watcher1")
        self.frame_7 = self.findChild(QFrame, "frame_7")
        self.lineEdit_Inc3 = self.findChild(QLineEdit, "lineEdit_Inc3")
        self.lineEdit_Inc1 = self.findChild(QLineEdit, "lineEdit_Inc1")
        self.label_8 = self.findChild(QLabel, "label_8")
        self.lineEdit_Inc2 = self.findChild(QLineEdit, "lineEdit_Inc2")
        self.label_6 = self.findChild(QLabel, "label_6")
        self.label_7 = self.findChild(QLabel, "label_7")
        self.lineEdit_inc_workshop = self.findChild(QLineEdit, "lineEdit_inc_workshop")
        self.lineEdit_inc_prof = self.findChild(QLineEdit, "lineEdit_inc_prof")
        self.frame_8 = self.findChild(QFrame, "frame_8")
        self.label_15 = self.findChild(QLabel, "label_15")
        self.LogFrame = self.findChild(QTextEdit, "LogFrame")
        self.ProgressBar = self.findChild(QProgressBar, "ProgressBar")
        self.StartButton = self.findChild(QPushButton, "StartButton")
        self.CounterOfStarts = self.findChild(QLCDNumber, "CounterOfStarts")
        self.frame_9 = self.findChild(QFrame, "frame_9")
        self.label_10 = self.findChild(QLabel, "label_10")
        self.label_12 = self.findChild(QLabel, "label_12")
        self.label_11 = self.findChild(QLabel, "label_11")
        self.spinBox_2 = self.findChild(QSpinBox, "spinBox_2")
        self.spinBox = self.findChild(QSpinBox, "spinBox")
        self.frame_10 = self.findChild(QFrame, "frame_10")
        self.label_13 = self.findChild(QLabel, "label_13")
        self.lineEdit_Name = self.findChild(QLineEdit, "lineEdit_Name")
        self.StopButton = self.findChild(QPushButton, "StopButton")
        self.menubar = self.findChild(QMenuBar, "menubar")
        self.statusbar = self.findChild(QStatusBar, "statusbar")
        self.label_16 = self.findChild(QLabel, "label_16")
        self.pushButton_Save_2 = self.findChild(QPushButton, "pushButton_Save_2")
        self.label_9 = self.findChild(QLabel, "label_9")
        self.Check_button = self.findChild(QPushButton, "Check_button")
        # Переменные для трёх исключений (текстовые данные в полях)
        self.name_inc1 = self.lineEdit_Inc1.text()
        self.name_inc2 = self.lineEdit_Inc2.text()
        self.name_inc3 = self.lineEdit_Inc3.text()
        # Переменные для цеха и кода профессии исключений (текстовые данные в полях)
        self.workshop_inc = self.lineEdit_inc_workshop.text()
        self.prof_inc = self.lineEdit_inc_prof.text()

        # Define Variables
        self.info_textbox_operations = ""
        self.info_textbox_save = ""
        self.info_textbox_template = ""
        self.info_textbox_tabels = ""
        self.info_textbox_cal = ""

        # Do something
        self.pushButton_Marsh.clicked.connect(self.clickbutton_operations)
        self.pushButton_Save.clicked.connect(self.clickbutton_save)
        self.pushButton_Main.clicked.connect(self.clickbutton_template)
        self.pushButton_Cal.clicked.connect(self.clickbutton_cal)
        self.StartButton.clicked.connect(self.clickbutton_start)
        self.StopButton.clicked.connect(self.clickbutton_stop)
        self.pushButton_Save_2.clicked.connect(self.get_space_tabels)
        self.Check_button.clicked.connect(self.get_check_data)

        # Set color
        self.label.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )
        self.label_4.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )
        self.label_5.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )
        self.label_16.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )
        self.label_9.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )
        self.label_13.setStyleSheet(
            "color: rgb(255, 255, 255); background-color: rgba(25, 10, 25, 100);"
            'font: 87 8pt "Arial Black"; border-radius: 7px;'
        )

        # Установка фона
        self.setStyleSheet(
            "#MainWindow{border-image:url(Chronos_background.jpg);"
            "background-repeat: no-repeat;}"
        )
        self.themes_box.addItems(
            [
                "Chronos_background.jpg",
                "background_img1.jpg",
                "background_img2.jpg",
                "background_img3.jpg",
                "background_img4.jpg",
            ]
        )
        self.themes_box.currentIndexChanged.connect(self.set_themes)

        # Show app
        self.show()

    # Setting themes
    def set_themes(self, index):
        selectedtheme = self.themes_box.itemText(index)
        stylesheet = f"#MainWindow {{border-image: url({selectedtheme}); background-repeat: no-repeat;}}"
        self.setStyleSheet(stylesheet)

    # Загрузить файл операций и изменить цвет лейбла
    def clickbutton_operations(self):
        self.info_textbox_operations = str(QFileDialog.getOpenFileName(self)[0])
        self.LogFrame.insertPlainText(
            f"\nВыбран файл операций: {self.info_textbox_operations}"
        )
        self.label.setStyleSheet(
            "color: rgb(255, 255, 255);"
            "background-color: rgba(255, 255, 255, 0);"
            'font: 87 8pt "Arial Black";'
        )
        self.wb_operations = opxl.load_workbook(self.info_textbox_operations)

    # Загрузить путь сохранения и изменить цвет лейбла
    def clickbutton_save(self):
        self.info_textbox_save = str(QFileDialog.getExistingDirectory(self))
        self.LogFrame.insertPlainText(
            "\nВыбран путь сохранения: " + str(self.info_textbox_save)
        )
        self.label_4.setStyleSheet(
            "color: rgb(255, 255, 255);"
            "background-color: rgba(255, 255, 255, 0);"
            'font: 87 8pt "Arial Black";'
        )

    # Загрузить файл-шаблон и изменить цвет лейбла
    def clickbutton_template(self):
        self.info_textbox_template = str(QFileDialog.getOpenFileName(self)[0])
        self.LogFrame.insertPlainText(
            "\nВыбран шаблон: " + str(self.info_textbox_template)
        )
        self.label_5.setStyleSheet(
            "color: rgb(255, 255, 255);"
            "background-color: rgba(255, 255, 255, 0);"
            'font: 87 8pt "Arial Black";'
        )
        self.wb_template = opxl.load_workbook(self.info_textbox_template)

    # Загрузить файл-календарь и изменить цвет лейбла
    def clickbutton_cal(self):
        self.info_textbox_cal = str(QFileDialog.getOpenFileName(self)[0])
        self.LogFrame.insertPlainText(
            f"\nВыбран календарь: {str(self.info_textbox_cal)}"
        )
        self.label_9.setStyleSheet(
            "color: rgb(255, 255, 255);"
            "background-color: rgba(255, 255, 255, 0);"
            'font: 87 8pt "Arial Black";'
        )
        self.wb_calendar = opxl.load_workbook(self.info_textbox_cal)

    # Загрузить путь к папке табелей и изменить цвет лейбла
    def get_space_tabels(self):
        self.info_textbox_tabels = str(QFileDialog.getExistingDirectory(self))
        self.LogFrame.insertPlainText(
            "\nВыбрана папка табелей: " + str(self.info_textbox_tabels)
        )
        self.label_16.setStyleSheet(
            "color: rgb(255, 255, 255);"
            "background-color: rgba(255, 255, 255, 0);"
            'font: 87 8pt "Arial Black";'
        )

    def get_check_data(self):
        if any(
            value == ""
            for value in [self.info_textbox_operations, self.info_textbox_tabels]
        ):
            self.LogFrame.insertPlainText(
                "\nПроверка недоступна! Выберите файл операций и папку табелей."
            )
        else:
            self.core_Object.get_random_tabel()
            combinations = set()
            combinations_2 = set()
            for row in self.core_Object.wb_tabel[
                self.core_Object.wb_tabel.sheetnames[0]
            ].iter_rows(min_row=2, values_only=True):
                workshop, cipher = row[0], row[8]
                combination = (workshop, cipher)
                combinations.add(combination)
                workshop = row[1]
                try:
                    int(workshop)
                    combination_2 = (workshop, cipher)
                    combinations_2.add(combination_2)
                except:
                    pass
            missing_combinations = set()
            for row in self.wb_operations[self.wb_operations.sheetnames[0]].iter_rows(
                min_row=2, values_only=True
            ):
                workshop, cipher = (
                    row[1],
                    row[3],
                )
                if isinstance(workshop, int):
                    combination = (workshop, cipher)
                    if combination not in combinations and combinations_2:
                        missing_combinations.add(combination)
            if missing_combinations:
                self.LogFrame.insertPlainText(f"\nНет в табеле:")
                for combination in missing_combinations:
                    self.LogFrame.insertPlainText(f"\n{combination}")
            else:
                self.LogFrame.insertPlainText(f"\nВсё совпадает!")

    # Сохранить и записать в файл Log, и выйти из приложения
    def clickbutton_stop(self):
        with open("Logs.txt", "w+", encoding="utf-8") as log_file:
            log_file.write(f"{self.LogFrame.toPlainText()}")
            log_file.close()
        sys.exit()

    # Инициализация функций и так далее по нажатию на кнопку "Пуск"
    def clickbutton_start(self):
        # Переменная для наименования изделия (текст)
        self.name_product = self.lineEdit_Name.text()
        # Проверки на заполненность полей и изменение цвета виджетов текста
        if any(
            value == ""
            for value in [
                self.info_textbox_operations,
                self.info_textbox_save,
                self.name_product,
                self.info_textbox_template,
                self.info_textbox_tabels,
                self.info_textbox_cal,
            ]
        ):
            self.LogFrame.insertPlainText("\nВыбраны не все элементы!")
            if self.info_textbox_operations == "":
                self.label.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )
            if self.info_textbox_save == "":
                self.label_4.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )
            if self.info_textbox_template == "":
                self.label_5.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )
            if self.name_product == "":
                self.label_13.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )
            if self.info_textbox_tabels == "":
                self.label_16.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )
            if self.info_textbox_cal == "":
                self.label_9.setStyleSheet(
                    "color: rgb(255, 255, 255);"
                    "background-color: rgba(255, 50, 50, 100);"
                    'font: 87 8pt "Arial Black"; border-radius: 7px;'
                )

        else:  # Печатаем в лог информацию о изделии
            self.LogFrame.insertPlainText(
                "\n----------------------------------------------"
                "\nИзделие: " + self.name_product
            )
            self.LogFrame.insertPlainText(
                "\nМесяцы: с "
                + str(self.spinBox.value())
                + " по "
                + str(self.spinBox_2.value())
            )

            if (
                (self.name_inc1 != "")
                or (self.name_inc2 != "")
                or (self.name_inc3 != "")
            ):
                self.LogFrame.insertPlainText(
                    "\nИсключения: "
                    + self.name_inc1
                    + ", "
                    + self.name_inc2
                    + ", "
                    + self.name_inc3
                    + ", "
                    + "\nЦех (исключение): "
                    + self.workshop_inc
                    + "\nКод профессии (исключение): "
                    + self.prof_inc
                )
            else:
                self.LogFrame.insertPlainText("\nНет исключений")

            # Adds

            self.LogFrame.insertPlainText(
                f"\n{self.core_Object.get_count_of_space()}----Всего строк в файле операций"
            )
            self.ProgressBar.setValue(25)
            self.LogFrame.insertPlainText(
                f"\n{self.core_Object.get_way_rows()}----Количество деталей"
            )
            self.LogFrame.insertPlainText(
                f"\nСписок стартовых строк => {self.core_Object.list_startrows_of_things}"
            )

            self.core_Object.make_buty()
            self.ProgressBar.setValue(100)


class Core:
    def __init__(self, interface_Object):
        self.interface_Object = interface_Object

    # Генератор случайных чисел с отклонением в 20 процентов
    def generate_values(self):
        self.values = [
            round(random.uniform(0.9 * self.target_mean, 1.1 * self.target_mean), 5)
            for _ in range(3)
        ]
        # Текущее среднее
        current_mean = sum(self.values) / 3
        # Масштабирум значения, чтобы среднее значение было равно целевому
        scaled_values = [
            round(value * self.target_mean / current_mean, 5) for value in self.values
        ]
        return scaled_values

    # Вспомогательная функция к initialize_exel, которая генерирует строку чисел по правилам и перемешивает её
    def initialize_exel_2(self):
        current = random.uniform(1.03, 1.09)
        percents = [1, current]

        for _ in range(8 - 2):
            min_border = max(1 / current, 0.9)
            max_border = min(percents[1] / current, 1.1)

            current *= random.uniform(min_border, max_border)
            percents.append(current)

        coeff = self.middle / sum(percents) * len(percents)
        result = [round(percent * coeff, 5) for percent in percents]
        self.coef_chrono = round(max(result) / min(result), 2)
        y = random.randint(972, 990) / 1000
        minimal_ghost_value = round(min(result) * y, 5)
        print(minimal_ghost_value, "<-Минимальное гостовое")
        a = (random.randint(11150, 15988) / 100000) + 1
        x = min(result) * y * a / max(result)
        maximal_ghost_value = round(max(result) * x, 5)
        print(maximal_ghost_value, "<-Максимальное гостовое")
        print("Высокий множитель:", x, "Низкий множитель:", y)
        print(a, "<-Заданный предел")
        print(round(maximal_ghost_value / minimal_ghost_value, 2))
        result.append(minimal_ghost_value)
        result.append(maximal_ghost_value)
        random.shuffle(result)
        return result

    # Заполняет числовые поля относительно target_mean(Топер) в шаблон
    def initialize_exel(self):
        self.target_mean = 0.25694  #!_Приписывать итерративно
        temporary_massiv = []
        temporary_massiv = self.generate_values()
        counter = 0
        for i in range(15, 22, 3):
            self.interface_Object.wb_template[
                self.interface_Object.wb_template.sheetnames[0]
            ].cell(row=i, column=14, value=round(temporary_massiv[counter], 5))
            temporary_massiv_2 = []
            self.target_mean = float(temporary_massiv[counter])
            self.middle = temporary_massiv[counter]
            temporary_massiv_2 = self.initialize_exel_2()
            self.interface_Object.wb_template[
                self.interface_Object.wb_template.sheetnames[0]
            ].cell(row=i, column=15, value=self.coef_chrono)
            for j in range(1, 11):
                self.interface_Object.wb_template[
                    self.interface_Object.wb_template.sheetnames[0]
                ].cell(row=i, column=j + 3, value=temporary_massiv_2[j - 1])
            counter += 1

    # Возвращает общее количество строк в файле операций
    def get_count_of_space(self):
        total = self.interface_Object.wb_operations[
            self.interface_Object.wb_operations.sheetnames[0]
        ].max_row
        return total

    # Возвращает количество деталей (считает строки начинающиеся с "Наименование") в файле операций
    def get_way_rows(self):
        self.NamesOfThings = []
        self.list_startrows_of_things = []
        s = 0
        j = self.get_count_of_space()
        for i in range(1, j + 1):
            if str(
                self.interface_Object.wb_operations[
                    self.interface_Object.wb_operations.sheetnames[0]
                ]
                .cell(row=i, column=2)
                .value
            ).startswith("Наименование"):
                s += 1
                self.NamesOfThings.append(i)
                self.list_startrows_of_things.append(
                    i + 3
                )  # i+3 это разница между наименованием и началом строк данных
        return s

    # Загружает в память и возвращает случайный табель из папки в интервале SpinBox-SpinBox_2
    def get_random_tabel(self):
        tabel = f"Tab{random.randint(self.interface_Object.spinBox.value(), self.interface_Object.spinBox_2.value())}.xlsx"
        self.wb_tabel = opxl.load_workbook(
            f"{self.interface_Object.info_textbox_tabels}/{tabel}"
        )
        return tabel

    def test_prog(self):
        self.interface_Object.LogFrame.insertPlainText(
            f"\nВремя выполнения: "
            f"{round((time.time() - self.interface_Object.timer), 2)}"
        )
        try:
            self.interface_Object.LogFrame.insertPlainText(
                f"\nПервый файл инициализирован:{self.interface_Object.info_textbox_operations}"
            )
        except AttributeError:
            self.interface_Object.LogFrame.insertPlainText(
                f"\nПеременные не были объявлены!"
            )

    # Возвращает номера колонок, которые соответствуют пустым ячейкам; Prime_rows - массив взятых людей
    def get_random_day(self, row_number):  # ТРЕБУЕТ ДОРАБОТКИ!
        Temp_rows = []
        for i in range(31, 93, 2):
            if (
                self.wb_tabel[self.wb_tabel.sheetnames[0]]
                .cell(row=row_number + 2, column=i)
                .value
                is not None
            ) and (
                self.wb_tabel[self.wb_tabel.sheetnames[0]]
                .cell(row=row_number + 2, column=i + 1)
                .value
                is None
            ):
                Temp_rows.append(i + 1)
        return Temp_rows

    def get_people_original(self):
        people = []
        # people_alt = []
        # cypher_alt = self.interface_Object.alternative_codes.get(self.cypher)
        for row_number, row in enumerate(
            self.wb_tabel[self.wb_tabel.sheetnames[0]].iter_rows(min_row=2)
        ):
            try:
                if (
                    (
                        row[0].value == self.workshop
                        or int(row[1].value) == int(self.workshop)
                    )
                    and (row[7].value == 11 or row[7].value == 12)
                    and row[8].value == self.cypher
                    and row[0].value != 250
                    and self.get_random_day(row_number)


                ):

                    people.append(row_number + 2)
            except:
                pass
        return people

    #  Цикл поиска людей по шифру в Острогоржске
    def get_people_Ostrog(self):
        people = []
        for row_number, row in enumerate(
            self.wb_tabel[self.wb_tabel.sheetnames[0]].iter_rows(min_row=2)
        ):
            if (
                row[0].value == 250
                and (row[7] == 11 or row[7] == 12)
                and row[8] == self.cypher
                and self.get_random_day(row_number)



            ):
                people.append(row_number + 2)
            elif (
                row[0].value == 250
                and (row[7] == 11 or row[7] == 12)
                and row[8] == random.choice(self.interface_Object.alternative_codes.get(self.cypher))  # 05052024 Продолжить
                and self.get_random_day(row_number)



            ):
                people.append(row_number + 2)

        return people

    def make_new_list(self):
        pass

    # Функция возвращает наименования относительно спика строк NamesOfThings (начинающиеся с "Наименование") и делает
    # дополнительные иттеративные действия
    def make_buty(self):
        self.initialize_exel()  # Добавить в цикл на каждый лист!!! len(self.NamesOfThings)
        # Новое изделие на новом листе
        for makes in range(4):
            Name_thing = str(
                self.interface_Object.wb_operations[
                    self.interface_Object.wb_operations.sheetnames[0]
                ]
                .cell(row=self.NamesOfThings[makes], column=2)
                .value
            )[46:]
            self.interface_Object.LogFrame.insertPlainText(
                f"\nНаименования =>{Name_thing}"
            )
            self.interface_Object.LogFrame.insertPlainText(
                f"\n{self.get_random_tabel()}<-Выбранный табель"
            )
            startcell = self.list_startrows_of_things[makes]
            # Цикл для заполнения и создания листов деталей в изделии
            while (
                self.interface_Object.wb_operations[
                    self.interface_Object.wb_operations.sheetnames[0]
                ]
                .cell(row=startcell, column=6)
                .value
                is not None
            ):
                if (
                    self.interface_Object.wb_operations[
                        self.interface_Object.wb_operations.sheetnames[0]
                    ]
                    .cell(row=startcell, column=6)
                    .value
                    != 0
                ):
                    # Присваиваю значения со строки (цех, операцию, шифр, разряд, Топер)
                    self.workshop = int(
                        self.interface_Object.wb_operations[
                            self.interface_Object.wb_operations.sheetnames[0]
                        ]
                        .cell(row=startcell, column=2)
                        .value
                    )
                    self.operation = (
                        self.interface_Object.wb_operations[
                            self.interface_Object.wb_operations.sheetnames[0]
                        ]
                        .cell(row=startcell, column=3)
                        .value
                    )
                    self.cypher = (
                        self.interface_Object.wb_operations[
                            self.interface_Object.wb_operations.sheetnames[0]
                        ]
                        .cell(row=startcell, column=4)
                        .value
                    )
                    self.work_number = (
                        self.interface_Object.wb_operations[
                            self.interface_Object.wb_operations.sheetnames[0]
                        ]
                        .cell(row=startcell, column=5)
                        .value
                    )
                    self.timing = (
                        self.interface_Object.wb_operations[
                            self.interface_Object.wb_operations.sheetnames[0]
                        ]
                        .cell(row=startcell, column=6)
                        .value
                    )

                    self.interface_Object.LogFrame.insertPlainText(
                        f"\nЯчейка: {startcell}, Цех: {self.workshop}, Операция: {self.operation}, Шифр: {self.cypher}, Разряд: {self.work_number}, Топер: {self.timing}"
                    )
                    if self.workshop == 250:
                        self.interface_Object.LogFrame.insertPlainText(
                            f"\nВыбранные люди => {self.get_people_Ostrog()}"
                        )
                    else:
                        self.interface_Object.LogFrame.insertPlainText(
                            f"\nВыбранные люди => {self.get_people_original()}"
                        )
                startcell += 1

            # Временно привязал lineEdit_Name, чтобы сохранять под разными именами
            # self.interface_Object.wb_template.save(
            # f"{self.interface_Object.info_textbox_save}/{Name_thing}.xlsx"
            # )


app = QApplication(sys.argv)  # Присваивание переменной действий приложений
UIWindow = UI()
app.exec_()
