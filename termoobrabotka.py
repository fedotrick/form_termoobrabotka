import sys
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QComboBox, QDateEdit, QHBoxLayout
)
from PySide6.QtCore import Qt, QDate
from openpyxl import Workbook, load_workbook
from PySide6.QtGui import QPalette, QColor, QFont

# Функция для сохранения данных в Excel
def save_to_excel(номер_плавки, термообработка_номер_печи, термообработка_дата,
                 термообработка_начало_первого_цикла, термообработка_конец_первого_цикла,
                 термообработка_начало_второго_цикла="", термообработка_конец_второго_цикла=""):
    try:
        # Если файл не существует, создаем его с заголовками
        if not os.path.exists('termoobrabotka.xlsx'):
            wb = Workbook()
            ws = wb.active
            ws.title = "Records"
            headers = ['Номер плавки', 'Номер печи', 'Дата', 
                      'Начало первого цикла', 'Конец первого цикла',
                      'Начало второго цикла', 'Конец второго цикла']
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header)
        else:
            wb = load_workbook('termoobrabotka.xlsx')
            if "Records" not in wb.sheetnames:
                ws = wb.create_sheet("Records")
                headers = ['Номер плавки', 'Номер печи', 'Дата', 
                          'Начало первого цикла', 'Конец первого цикла',
                          'Начало второго цикла', 'Конец второго цикла']
                for col, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=col, value=header)
            else:
                ws = wb["Records"]
            
        next_row = ws.max_row + 1
        
        values = [номер_плавки, термообработка_номер_печи, термообработка_дата,
                 термообработка_начало_первого_цикла, термообработка_конец_первого_цикла,
                 термообработка_начало_второго_цикла, термообработка_конец_второго_цикла]
        
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=next_row, column=col)
            cell.value = value
            
            if col == 3:  # Колонка C (дата)
                cell.number_format = 'DD.MM.YYYY'
            elif col in [4, 5, 6, 7]:  # Колонки D, E, F, G (время)
                cell.number_format = 'HH:MM'
        
        wb.save('termoobrabotka.xlsx')
        wb.close()
    except Exception as e:
        raise Exception(f"Ошибка при сохранении в Excel: {str(e)}")

def get_existing_plavki():
    file_name = 'plavka.xlsx'
    if not os.path.exists(file_name):
        return []

    workbook = load_workbook(file_name)
    sheet = workbook.active
    номера_плавок = []
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        номер_плавки = row[1]  
        if номер_плавки is not None:
            номера_плавок.append(str(номер_плавки))

    return номера_плавок

def get_available_plavki():
    # Получаем все номера плавок из plavka.xlsx
    все_плавки = get_existing_plavki()
    
    # Получаем номера плавок, которые уже существуют в termoobrabotka.xlsx
    существующие_плавки = set()
    file_name = 'termoobrabotka.xlsx'
    if os.path.exists(file_name):
        workbook = load_workbook(file_name)
        if "Records" in workbook.sheetnames:
            sheet = workbook["Records"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                существующие_плавки.add(row[0])  # Первый столбец содержит номер плавки

    # Фильтруем номера плавок, оставляя только те, которые отсутствуют в termoobrabotka.xlsx и имеют "/25"
    доступные_плавки = [
        плавка for плавка in все_плавки 
        if плавка not in существующие_плавки and '/25' in плавка
    ]

    # Сортируем по убыванию
    доступные_плавки.sort(reverse=True)

    # Отладочные сообщения
    print(f"Все плавки: {все_плавки}")
    print(f"Существующие плавки: {существующие_плавки}")
    print(f"Доступные плавки: {доступные_плавки}")

    return доступные_плавки


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.plавка_fields = []
        
        self.setWindowTitle("Электронный журнал термообработки")
        self.setMinimumWidth(600)  # Уменьшаем минимальную ширину
        self.setMinimumHeight(400)  # Уменьшаем минимальную высоту
        
        # Обновляем стили для более компактного вида
        self.setStyleSheet("""
            QWidget {
                background-color: #1a1a1a;
                color: #ffffff;
                font-family: 'Arial';
            }
            QLabel {
                color: #00ffff;
                font-size: 12px;  /* Уменьшаем размер шрифта */
                padding: 5px;     /* Уменьшаем отступы */
                border: 1px solid #00ffff;
                border-radius: 3px;
                background-color: #2a2a2a;
                margin: 1px;      /* Добавляем минимальные внешние отступы */
            }
            QComboBox {
                background-color: #2a2a2a;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 3px;
                color: #ffffff;
                min-height: 20px; /* Уменьшаем минимальную высоту */
                margin: 1px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QLineEdit {
                background-color: #2a2a2a;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 3px;
                color: #ffffff;
                min-height: 20px;
                margin: 1px;
            }
            QDateEdit {
                background-color: #2a2a2a;
                border: 1px solid #00ffff;
                border-radius: 3px;
                padding: 3px;
                color: #ffffff;
                min-height: 20px;
                margin: 1px;
            }
            QPushButton {
                background-color: #00ffff;
                color: #000000;
                border: none;
                border-radius: 3px;
                padding: 5px;
                font-size: 12px;
                font-weight: bold;
                min-height: 25px;
                margin: 1px;
            }
            QPushButton:hover {
                background-color: #00cccc;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(2)  # Уменьшаем расстояние между элементами
        layout.setContentsMargins(5, 5, 5, 5)  # Уменьшаем отступы от краёв

        # Заголовок
        title = QLabel("ЭЛЕКТРОННЫЙ ЖУРНАЛ ТЕРМООБРАБОТКИ")
        title.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #00ffff;
                border: 2px solid #00ffff;
                padding: 5px;
                background-color: #2a2a2a;
            }
        """)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Создаем горизонтальные контейнеры для группировки элементов
        top_container = QWidget()
        top_layout = QHBoxLayout(top_container)
        top_layout.setSpacing(2)
        top_layout.setContentsMargins(0, 0, 0, 0)

        # Левая часть с печью и плавками
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setSpacing(2)
        left_layout.setContentsMargins(0, 0, 0, 0)

        печь_label = QLabel("НОМЕР ПЕЧИ")
        печь_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(печь_label)

        self.термообработка_номер_печи = QComboBox()
        self.термообработка_номер_печи.addItems(['1', '2'])
        self.термообработка_номер_печи.currentTextChanged.connect(self.update_plavka_fields)
        left_layout.addWidget(self.термообработка_номер_печи)

        # Контейнер для плавок
        self.plавка_container = QWidget()
        self.plавка_layout = QVBoxLayout(self.plавка_container)
        self.plавка_layout.setSpacing(2)
        self.plавка_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(self.plавка_container)

        top_layout.addWidget(left_container)

        # Правая часть с датой и циклами
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(2)
        right_layout.setContentsMargins(0, 0, 0, 0)

        дата_label = QLabel("ТЕРМООБРАБОТКА")
        дата_label.setAlignment(Qt.AlignCenter)
        дата_label.setStyleSheet("""
            QLabel {
                color: #000000;
                font-size: 32px;
                font-weight: bold;
                padding: 5px;
                border: 1px solid #00ffff;
                border-radius: 3px;
                background-color: #2a2a2a;
                margin: 1px;
                background-image: url(termo.png);
                background-position: center;
                background-repeat: no-repeat;
                background-origin: content;
            }
        """)
        right_layout.addWidget(дата_label)

        self.термообработка_дата = QDateEdit()
        self.термообработка_дата.setDisplayFormat("dd.MM.yyyy")
        self.термообработка_дата.setCalendarPopup(True)
        self.термообработка_дата.setDate(QDate.currentDate())
        right_layout.addWidget(self.термообработка_дата)

        # Контейнер для циклов
        cycles_container = QWidget()
        cycles_layout = QHBoxLayout(cycles_container)
        cycles_layout.setSpacing(2)
        cycles_layout.setContentsMargins(0, 0, 0, 0)

        # Первый цикл
        cycle1_container = QWidget()
        cycle1_layout = QVBoxLayout(cycle1_container)
        cycle1_layout.setSpacing(2)
        cycle1_layout.setContentsMargins(0, 0, 0, 0)

        цикл1_label = QLabel("ПЕРВЫЙ ЦИКЛ")
        цикл1_label.setAlignment(Qt.AlignCenter)
        cycle1_layout.addWidget(цикл1_label)

        self.термообработка_начало_первого_цикла = QLineEdit()
        self.термообработка_начало_первого_цикла.setPlaceholderText("Начало (ЧЧ:ММ)")
        self.термообработка_начало_первого_цикла.setMaxLength(5)
        self.термообработка_начало_первого_цикла.textChanged.connect(self.format_time_input)
        cycle1_layout.addWidget(self.термообработка_начало_первого_цикла)

        self.термообработка_конец_первого_цикла = QLineEdit()
        self.термообработка_конец_первого_цикла.setPlaceholderText("Конец (ЧЧ:ММ)")
        self.термообработка_конец_первого_цикла.setMaxLength(5)
        self.термообработка_конец_первого_цикла.textChanged.connect(self.format_time_input)
        cycle1_layout.addWidget(self.термообработка_конец_первого_цикла)

        cycles_layout.addWidget(cycle1_container)

        # Второй цикл
        cycle2_container = QWidget()
        cycle2_layout = QVBoxLayout(cycle2_container)
        cycle2_layout.setSpacing(2)
        cycle2_layout.setContentsMargins(0, 0, 0, 0)

        цикл2_label = QLabel("ВТОРОЙ ЦИКЛ")
        цикл2_label.setAlignment(Qt.AlignCenter)
        cycle2_layout.addWidget(цикл2_label)

        self.термообработка_начало_второго_цикла = QLineEdit()
        self.термообработка_начало_второго_цикла.setPlaceholderText("Начало (ЧЧ:ММ)")
        self.термообработка_начало_второго_цикла.setMaxLength(5)
        self.термообработка_начало_второго_цикла.textChanged.connect(self.format_time_input)
        cycle2_layout.addWidget(self.термообработка_начало_второго_цикла)

        self.термообработка_конец_второго_цикла = QLineEdit()
        self.термообработка_конец_второго_цикла.setPlaceholderText("Конец (ЧЧ:ММ)")
        self.термообработка_конец_второго_цикла.setMaxLength(5)
        self.термообработка_конец_второго_цикла.textChanged.connect(self.format_time_input)
        cycle2_layout.addWidget(self.термообработка_конец_второго_цикла)

        cycles_layout.addWidget(cycle2_container)
        right_layout.addWidget(cycles_container)

        top_layout.addWidget(right_container)
        layout.addWidget(top_container)

        # Кнопка сохранения
        self.save_button = QPushButton("СОХРАНИТЬ")
        self.save_button.clicked.connect(self.save_data)
        layout.addWidget(self.save_button)

        self.setLayout(layout)
        
        # Инициализация полей плавок
        self.update_plavka_fields('1')

    def update_plavka_fields(self, печь_номер):
        # Очищаем существующие поля
        for field in self.plавка_fields:
            self.plавка_layout.removeWidget(field)
            field.deleteLater()
        self.plавка_fields.clear()

        # Определяем количество полей в зависимости от номера печи
        количество_полей = 10 if печь_номер == '1' else 9
        
        # Получаем доступные плавки
        доступные_плавки = get_available_plavki()
        
        # Создаем новые поля
        for i in range(количество_полей):
            combo = QComboBox()
            combo.addItem(f"ПЛАВКА {i+1}")
            combo.addItems(доступные_плавки)
            self.plавка_fields.append(combo)
            self.plавка_layout.addWidget(combo)

    def format_time_input(self, text):
        """Автоматически добавляет двоеточие после двух цифр"""
        if len(text) == 2 and text.isdigit():
            self.sender().setText(text + ":")
            self.sender().setCursorPosition(3)  # Установка курсора после двоеточия

    def validate_time(self, time_str):
        """Проверка корректности ввода времени в формате ЧЧ:ММ"""
        try:
            hours, minutes = map(int, time_str.split(':'))
            if 0 <= hours < 24 and 0 <= minutes < 60:
                return True
        except ValueError:
            return False
        return False

    def save_data(self):
        номер_печи = self.термообработка_номер_печи.currentText()
        термообработка_дата = self.термообработка_дата.date().toString("dd.MM.yyyy")
        начало_первого_цикла = self.термообработка_начало_первого_цикла.text().strip()
        конец_первого_цикла = self.термообработка_конец_первого_цикла.text().strip()
        начало_второго_цикла = self.термообработка_начало_второго_цикла.text().strip()
        конец_второго_цикла = self.термообработка_конец_второго_цикла.text().strip()

        # Проверка времени для первого цикла
        if not self.validate_time(начало_первого_цикла) or not self.validate_time(конец_первого_цикла):
            QMessageBox.warning(self, "Ошибка", "Некорректный ввод времени первого цикла. Используйте формат ЧЧ:ММ.")
            return

        # Проверка времени для второго цикла (если заполнено)
        if (начало_второго_цикла or конец_второго_цикла):
            if not self.validate_time(начало_второго_цикла) or not self.validate_time(конец_второго_цикла):
                QMessageBox.warning(self, "Ошибка", "Некорректный ввод времени второго цикла. Используйте формат ЧЧ:ММ.")
                return

        # Собираем выбранные плавки
        выбранные_плавки = []
        for combo in self.plавка_fields:
            плавка = combo.currentText()
            if плавка != f"ПЛАВКА {len(выбранные_плавки)+1}":
                выбранные_плавки.append(плавка)

        if not выбранные_плавки:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну плавку.")
            return

        try:
            # Сохраняем данные для каждой выбранной плавки
            for плавка in выбранные_плавки:
                save_to_excel(
                    плавка,
                    номер_печи,
                    термообработка_дата,
                    начало_первого_цикла,
                    конец_первого_цикла,
                    начало_второго_цикла,
                    конец_второго_цикла
                )

            QMessageBox.information(self, "Успех", "Данные сохранены в Excel!")
            self.clear_fields()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении данных: {str(e)}")

    def clear_fields(self):
        self.термообработка_дата.setDate(QDate.currentDate())
        self.термообработка_начало_первого_цикла.clear()
        self.термообработка_конец_первого_цикла.clear()
        self.термообработка_начало_второго_цикла.clear()
        self.термообработка_конец_второго_цикла.clear()
        self.update_plavka_fields(self.термообработка_номер_печи.currentText())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
