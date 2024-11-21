import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class FileChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path.endswith("data.xlsx"):
            print("Файл data.xlsx изменён. Обновляем Google Таблицу...")
            update_google_sheet()


columns = ['ФИО', 'Должность', 'ДВ', 'Институт', 'Группа', 'Кафедра', 'Шифр', 
           'Телеграмм', 'ВК', 'Телефон', 'Почта', 'Отдел', 'День рождения']

def update_google_sheet():
    try:
        excel_file = 'data.xlsx'  # Имя файла с данными
        df = pd.read_excel(excel_file)  # Чтение данных в DataFrame
        df = df.replace([float('inf'), float('-inf'), float('nan')], None)
        df = df.fillna("")  # Заполнить NaN пустыми строками

        # Настройка Google Sheets API
        scope = ["https://spreadsheets.google.com/feeds", 
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file", 
                "https://www.googleapis.com/auth/drive"]
        
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)

        # Открытие нужной таблицы и запись данных
        sheet = client.open('Студенты').sheet1  # Открытие первого листа
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
        print("Данные успешно обновлены в Google Таблице!")
    except Exception as e:
        print(f"Ошибка при обновлении таблицы: {e}")


if __name__ == "__main__":
    event_handler = FileChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, path=".", recursive=False)
    print("Наблюдение за файлом data.xlsx запущено.")
    try:
        observer.start()
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


