import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Создание временного файла credentials.json из переменной окружения
def create_temp_credentials_file():
    credentials_content = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not credentials_content:
        raise ValueError("Переменная GOOGLE_APPLICATION_CREDENTIALS не задана!")
    
    temp_credentials_path = "/tmp/credentials.json"
    with open(temp_credentials_path, "w") as temp_file:
        temp_file.write(credentials_content)
    
    return temp_credentials_path

# Инициализация Google Sheets API
def initialize_google_client():
    credentials_path = create_temp_credentials_file()
    credentials = Credentials.from_service_account_file(credentials_path)
    return gspread.authorize(credentials)

# Функция для обновления Google Sheet
def update_google_sheet(file_path, sheet_name, worksheet_index=0):
    try:
        # Чтение Excel-файла
        df = pd.read_excel(file_path)

        # Авторизация Google Sheets
        client = initialize_google_client()

        # Открытие таблицы и выбор листа
        sheet = client.open(sheet_name).get_worksheet(worksheet_index)

        # Очистка листа
        sheet.clear()

        # Запись данных из DataFrame в Google Sheets
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
        print(f"Таблица '{sheet_name}' успешно обновлена!")
    except Exception as e:
        print(f"Ошибка обновления Google Sheets: {e}")

# Класс для обработки изменений в файле
class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def on_modified(self, event):
        if event.src_path == self.file_path:
            print(f"Изменения обнаружены в {self.file_path}. Обновляю Google Sheets...")
            update_google_sheet(self.file_path, self.sheet_name)

# Главный код
if __name__ == "__main__":
    # Путь к локальному Excel-файлу
    excel_file_path = "data.xlsx"
    google_sheet_name = "Студенты"

    # Проверка, существует ли файл
    if not os.path.exists(excel_file_path):
        print(f"Файл {excel_file_path} не найден!")
        exit(1)

    # Настройка наблюдения за изменениями
    event_handler = FileChangeHandler(excel_file_path, google_sheet_name)
    observer = Observer()
    observer.schedule(event_handler, path=os.path.dirname(excel_file_path), recursive=False)
    observer.start()

    print(f"Наблюдение за изменениями в файле {excel_file_path} запущено...")
    try:
        while True:
            pass  # Основной цикл, пока приложение работает
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
