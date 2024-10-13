import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import pandas as pd
import re
import threading
import queue  # Импортируем queue

# Переменная для фона
background_file = r"C:\Users\smirn\OneDrive\Рабочий стол\vscode\q2.jpg"  # Укажите путь к вашему изображению

# Создаем очередь для передачи сообщений
result_queue = queue.Queue()

# Функция для поиска файла
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

# Оптимизированная функция для поиска в Excel
def search_in_excel(filename, search_word):
    try:
        # Читаем все листы из Excel файла
        excel_data = pd.read_excel(filename, sheet_name=None)

        # Проверка: если данные пустые
        if not excel_data:
            raise ValueError("Файл Excel пуст или содержит некорректные данные.")

        # Словарь для хранения результатов
        results = {}

        # Подготовка шаблона регулярного выражения для поиска
        pattern = re.compile(fr'(^|\s){re.escape(search_word)}(\s|$)', re.IGNORECASE)

        # Проходим по каждому листу
        for sheet_name, data in excel_data.items():
            try:
                # Преобразуем все данные в строки, чтобы избежать ошибок
                data_str = data.astype(str)

                # Проверка каждого элемента строки на соответствие паттерну
                matches = data_str.applymap(lambda value: bool(pattern.search(value)))

                # Фильтруем строки, которые содержат искомое слово
                matched_rows = data[matches.any(axis=1)]

                # Если найдены строки, добавляем их в результаты
                if not matched_rows.empty:
                    results[sheet_name] = matched_rows

            except Exception as e:
                print(f"Ошибка при обработке листа '{sheet_name}': {e}")
                continue

        # Возвращаем результаты
        return results

    except Exception as e:
        raise RuntimeError(f"Ошибка при работе с файлом: {e}")

# Функция для обновления метки с результатами (обновляется в основном потоке)
def update_result_label():
    try:
        while True:  # Цикл для получения всех сообщений из очереди
            result_text = result_queue.get_nowait()
            result_display.delete(1.0, tk.END)  # Очищаем текстовое поле перед обновлением
            result_display.insert(tk.END, result_text)  # Вставляем новый текст
    except queue.Empty:
        pass

# Функция, которая запускает поиск в отдельном потоке
def start_search():
    filename = entry_file_path.get()
    search_word = entry_search.get()
    
    if not filename or not search_word:
        messagebox.showwarning("Предупреждение", "Пожалуйста, укажите файл и слово для поиска.")
    else:
        # Очищаем результат перед новым поиском
        result_display.delete(1.0, tk.END)
        result_display.insert(tk.END, "Идет поиск...")
        
        # Запускаем поиск в отдельном потоке
        search_thread = threading.Thread(target=run_search, args=(filename, search_word))
        search_thread.start()

# Функция, которая запускается в потоке и обновляет результаты
def run_search(filename, search_word):
    try:
        results = search_in_excel(filename, search_word)

        # Формируем строку для отображения результата
        if results:
            result_text = f"Результаты поиска по слову '{search_word}':\n"
            for sheet_name, matched_rows in results.items():
                result_text += f"\nЛист '{sheet_name}':\n"
                result_text += matched_rows.to_string(index=False) + "\n\n"  # Убираем индексы
        else:
            result_text = f"Слово '{search_word}' не найдено в файле."

        # Добавляем результат в очередь
        result_queue.put(result_text)
    except Exception as e:
        result_queue.put(f"Не удалось выполнить поиск: {e}")

# Создание основного окна
root = tk.Tk()
root.title("Поиск по каталогу")
root.geometry("1920x1080")  # Окно Full HD

# Добавление фона к окну (если используется изображение)
if background_file:
    image = Image.open(background_file)
    image = image.resize((1920, 1080), Image.ANTIALIAS)  # Масштабируем изображение под окно
    bg_image = ImageTk.PhotoImage(image)
    background_label = tk.Label(root, image=bg_image)
    background_label.place(relwidth=1, relheight=1)

# Создаем фрейм для центровки всех элементов и добавления белого фона
frame = tk.Frame(root, bg="white", bd=10)
frame.place(relx=0.5, rely=0.4, anchor="center", width=1500, height=450)

# Увеличенные размеры шрифтов
font_large = ("Arial", 20)

# Метка и поле для ввода пути файла
label_file_path = tk.Label(frame, text="Путь файла:", font=font_large, bg="white")
label_file_path.grid(row=0, column=0, padx=20, pady=20, sticky="e")

entry_file_path = tk.Entry(frame, width=50, font=font_large)
entry_file_path.grid(row=0, column=1, padx=20, pady=20)

browse_button = tk.Button(frame, text="Обзор", command=browse_file, font=font_large, width=10)
browse_button.grid(row=0, column=2, padx=20, pady=20)

# Метка и поле для ввода поиска
label_search = tk.Label(frame, text="Что вы хотите найти?", font=font_large, bg="white")
label_search.grid(row=1, column=0, padx=20, pady=20, sticky="e")

entry_search = tk.Entry(frame, width=50, font=font_large)
entry_search.grid(row=1, column=1, padx=20, pady=20)

# Кнопка для поиска
search_button = tk.Button(frame, text="Искать", font=font_large, width=20, command=start_search)
search_button.grid(row=2, column=1, pady=40)

# Добавляем надпись в правом верхнем углу
top_label = tk.Label(root, text="Программа работает около 5 минут,\nпожалуйста не выключайте программу", 
                     font=("Arial", 16), bg="white", fg="red")
top_label.place(relx=1.0, rely=0.0, anchor="ne", x=-20, y=20)

# Создаем текстовое поле для отображения результатов с прокруткой
result_display = scrolledtext.ScrolledText(root, font=("Arial", 14), wrap=tk.WORD, bg="white")
result_display.place(relx=0.5, rely=0.7, anchor="center", width=1500, height=300)

# Функция для периодической проверки очереди
def check_queue():
    update_result_label()  # Обновляем метку с результатами
    root.after(100, check_queue)  # Проверяем очередь каждые 100 мс

# Запускаем периодическую проверку очереди
check_queue()

# Запуск главного цикла программы
root.mainloop()
