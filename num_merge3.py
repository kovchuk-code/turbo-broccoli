import pymupdf
import fitz  # PyMuPDF (импортируем под обоими именами для совместимости)
import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ========== Функция для объединения PDF ==========
def merge_pdfs_from_folder(folder_path, output_filename="merge.pdf"):
    """
    Объединяет все PDF‑файлы из указанной папки в один документ.
    """
    # Формируем полный путь для сохранения итогового файла в той же папке
    output_path = os.path.join(folder_path, output_filename)

    # Находим все PDF‑файлы в папке (игнорируем вложенные папки)
    pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))

    if not pdf_files:
        messagebox.showwarning("Предупреждение", f"В папке '{folder_path}' не найдено PDF‑файлов.")
        return False

    print(f"Найдено {len(pdf_files)} PDF‑файлов. Начинаю объединение...")

    # Создаём новый PDF‑документ
    merged_pdf = pymupdf.open()

    for pdf_path in pdf_files:
        try:
            # Открываем текущий PDF
            pdf_document = pymupdf.open(pdf_path)
            # Вставляем все страницы из текущего PDF в объединённый документ
            merged_pdf.insert_pdf(pdf_document)
            print(f"Добавлен: {os.path.basename(pdf_path)}")
            # Закрываем текущий PDF, чтобы освободить ресурсы
            pdf_document.close()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обработке файла {pdf_path}: {e}")
            merged_pdf.close()
            return False

    # Сохраняем объединённый PDF в ту же папку
    try:
        merged_pdf.save(output_path)
        merged_pdf.close()
        messagebox.showinfo("Успех", f"Объединение завершено. Файл сохранён как: {output_path}")
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")
        merged_pdf.close()
        return False


# ========== Функция для добавления номеров страниц ==========
# ========== Функция для добавления номеров страниц ==========
# ========== Функция для добавления номеров страниц ==========
def add_page_numbers(input_pdf, output_pdf, start_page=1, start_number=1, 
                    use_background=True, font_name="helv", font_size=12,
                    offset_x=0, offset_y=0):
    """
    Добавляет номер страницы в правый верхний угол каждой страницы PDF.
    
    Параметры:
    - input_pdf: путь к входному PDF-файлу
    - output_pdf: путь для сохранения результата
    - start_page: с какой страницы документа начинать нумерацию (1 - первая страница)
    - start_number: с какого номера начинать отсчет
    - use_background: добавлять ли белый фон за текстом
    - font_name: название шрифта (Arial, Times New Roman, Isocpeur)
    - font_size: размер шрифта в пунктах
    - offset_x: смещение по X от стандартной позиции (положительное - влево, отрицательное - вправо)
    - offset_y: смещение по Y от стандартной позиции (положительное - вниз, отрицательное - вверх)
    """
    # Открываем PDF
    doc = fitz.open(input_pdf)
    total_pages = len(doc)
    
    # Проверяем корректность параметров
    if start_page < 1:
        messagebox.showwarning("Предупреждение", "Страница начала нумерации не может быть меньше 1. Установлено значение 1.")
        start_page = 1
    
    if start_page > total_pages:
        messagebox.showerror("Ошибка", f"Страница начала нумерации ({start_page}) больше общего количества страниц ({total_pages})")
        doc.close()
        return False

    # Словарь соответствия названий шрифтов в интерфейсе и в PyMuPDF
    font_mapping = {
        "Arial": "helv",           # Helvetica (Arial)
        "Times New Roman": "times", # Times New Roman
        "Isocpeur": "cour"          # Courier (как замена Isocpeur)
    }
    
    # Получаем внутреннее название шрифта
    internal_font = font_mapping.get(font_name, "helv")
    
    print(f"Начинаю нумерацию:")
    print(f"  - Всего страниц в документе: {total_pages}")
    print(f"  - Начинаю нумерацию со страницы: {start_page}")
    print(f"  - Начальный номер: {start_number}")
    print(f"  - Фон за текстом: {'Да' if use_background else 'Нет'}")
    print(f"  - Шрифт: {font_name} (внутренний: {internal_font})")
    print(f"  - Размер шрифта: {font_size} pt")
    print(f"  - Смещение по X: {offset_x} pt")
    print(f"  - Смещение по Y: {offset_y} pt")
    
    # Переменная для отслеживания текущего номера
    current_number = start_number
    
    for page_num in range(total_pages):
        page = doc[page_num]
        
        # Проверяем, нужно ли нумеровать эту страницу
        if page_num + 1 >= start_page:
            print(f"  Страница {page_num + 1}: добавляю номер {current_number}")
            
            # Получаем размер страницы
            rect = page.rect
            page_width = rect.width
            page_height = rect.height

            # Номер страницы
            text = f"{current_number}"
            
            # Увеличиваем номер для следующей страницы
            current_number += 1

            # Отступ от края по умолчанию
            default_margin = 30
            
            # Узнаем ширину и высоту текста
            text_width = fitz.get_text_length(text, fontname=internal_font, fontsize=font_size)
            text_height = font_size * 1.2

            # Определяем позицию в зависимости от поворота страницы
            if page.rotation == 0:
                # Страница без поворота - нормальная ориентация
                # Номер в правом верхнем углу
                x = page_width - default_margin - text_width - offset_x
                y = default_margin + offset_y
                
                # Прямоугольник для фона (под текстом)
                rect_bg = fitz.Rect(x-1, y - text_height+4, x + text_width, y+2)
                
                # Вставляем текст без дополнительного поворота
                insert_rotate = 0
                insert_point = (x, y)
                
            elif page.rotation == 90:
                # Страница повернута на 90 градусов по часовой стрелке
                # В системе координат PDF после поворота, чтобы текст был в правом верхнем углу:
                # Базовая точка: (page_height - default_margin, default_margin)
                base_x = page_height - default_margin
                base_y = default_margin
                
                # Для повернутой страницы текст тоже повернут на 90°
                # С учетом смещений
                x = base_x - offset_y  # offset_y влияет на горизонталь в повернутой системе
                y = base_y + text_width + offset_x  # offset_x влияет на вертикаль
                
                # Прямоугольник для фона (для повернутого текста)
                rect_bg = fitz.Rect(x - text_height, y - text_width, x, y)
                
                # Точка вставки и поворот
                insert_point = (x, y)
                insert_rotate = 90
                
            elif page.rotation == 180:
                # Страница повернута на 180 градусов
                # Верхний правый угол теперь в левом нижнем углу исходной страницы
                x = default_margin + offset_x
                y = page_height - default_margin - text_height - offset_y
                
                # Прямоугольник для фона
                rect_bg = fitz.Rect(x, y, x + text_width, y + text_height)
                
                # Точка вставки и поворот
                insert_point = (x, y + text_height)
                insert_rotate = 180
                
            elif page.rotation == 270:
                # Страница повернута на 270 градусов (или -90)
                # Симметрично rotation == 90, но в другую сторону
                base_x = page_height - default_margin
                base_y = page_width - default_margin
                
                # Для поворота 270 градусов
                x = base_x - text_height + offset_y
                y = base_y - text_width - offset_x
                
                # Прямоугольник для фона
                #rect_bg = fitz.Rect(x, y, x + text_height, y + text_width)
                rect_bg = fitz.Rect(x+ text_height-2, y-2 , x + 2* text_height-4, y+ text_width)               

                # Точка вставки и поворот
                insert_point = (x + text_height, y)
                insert_rotate = 270
                
            else:
                # На случай неизвестного поворота
                x = page_width - default_margin - text_width - offset_x
                y = default_margin + offset_y
                rect_bg = fitz.Rect(x, y - text_height, x + text_width, y)
                insert_point = (x, y)
                insert_rotate = 0

            # Вставляем фон за текстом (белый прямоугольник), если нужно
            if use_background:
                page.draw_rect(rect_bg, color=(1, 1, 1), fill=(1, 1, 1), width=1)

            # Вставляем текст на страницу с соответствующим поворотом
            page.insert_text(
                point=insert_point,
                text=text,
                fontsize=font_size,
                fontname=internal_font,
                color=(0, 0, 0),
                rotate=insert_rotate
            )
        else:
            print(f"  Страница {page_num + 1}: пропускаю (без номера)")

    # Сохраняем результат
    doc.save(output_pdf)
    doc.close()
    
    print(f"Готово! Всего пронумеровано страниц: {current_number - start_number}")
    return True

# ========== Класс для вкладки объединения PDF ==========
class MergePDFTab:
    def __init__(self, parent):
        self.parent = parent
        
        # Создаём фрейм для вкладки
        self.frame = ttk.Frame(parent, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.create_widgets()
    
    def create_widgets(self):
        # Метка для выбора папки
        ttk.Label(self.frame, text="Выберите папку с PDF‑файлами:", 
                 font=('Arial', 10)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # Поле ввода пути к папке
        self.folder_entry = ttk.Entry(self.frame, width=60, font=('Arial', 10))
        self.folder_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        # Кнопка для выбора папки
        ttk.Button(self.frame, text="Обзор...", 
                  command=self.browse_folder).grid(row=1, column=1, padx=5, pady=5)
        
        # Метка для имени выходного файла
        ttk.Label(self.frame, text="Имя итогового PDF‑файла:", 
                 font=('Arial', 10)).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        # Поле ввода имени выходного файла
        self.output_entry = ttk.Entry(self.frame, width=60, font=('Arial', 10))
        self.output_entry.insert(0, "merge.pdf")
        self.output_entry.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        
        # Кнопка запуска объединения
        ttk.Button(self.frame, text="Объединить PDF", 
                  command=self.start_merge, style='Accent.TButton').grid(row=4, column=0, pady=20)
        
        # Настройка растяжения колонок
        self.frame.columnconfigure(0, weight=1)
    
    def browse_folder(self):
        """Открывает диалоговое окно для выбора папки."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder_path)
    
    def start_merge(self):
        """Запускает процесс объединения PDF‑файлов."""
        folder_path = self.folder_entry.get().strip()
        output_filename = self.output_entry.get().strip()

        if not folder_path:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите папку с PDF‑файлами.")
            return

        if not output_filename:
            messagebox.showwarning("Предупреждение", "Пожалуйста, укажите имя итогового PDF‑файла.")
            return
        
        # Добавляем .pdf, если не указано
        if not output_filename.lower().endswith('.pdf'):
            output_filename += '.pdf'

        # Запускаем объединение
        merge_pdfs_from_folder(folder_path, output_filename)


# ========== Класс для вкладки нумерации страниц ==========
class NumberPagesTab:
    def __init__(self, parent):
        self.parent = parent
        
        # Создаём фрейм для вкладки
        self.frame = ttk.Frame(parent, padding="10")
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        self.create_widgets()
    
    def create_widgets(self):
        # Создаем основной контейнер с прокруткой для удобства
        main_canvas = tk.Canvas(self.frame, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # === Блок выбора файлов ===
        file_frame = ttk.LabelFrame(scrollable_frame, text="Выбор файлов", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Поле для ввода пути к входному файлу
        ttk.Label(file_frame, text="Входной PDF-файл:", 
                 font=('Arial', 10)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.input_entry = ttk.Entry(file_frame, width=60, font=('Arial', 10))
        self.input_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        ttk.Button(file_frame, text="Обзор...", 
                  command=self.select_input_file).grid(row=0, column=2, padx=5, pady=10)

        # Поле для ввода пути к выходному файлу
        ttk.Label(file_frame, text="Выходной PDF-файл:", 
                 font=('Arial', 10)).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.output_entry = ttk.Entry(file_frame, width=60, font=('Arial', 10))
        self.output_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        
        # Фрейм для кнопок выбора выходного файла
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=1, column=2, padx=5, pady=10)
        
        ttk.Button(button_frame, text="Выбрать папку...", 
                  command=self.select_output_folder).pack(side=tk.TOP, pady=2)
        ttk.Button(button_frame, text="Сохранить как...", 
                  command=self.select_output_file).pack(side=tk.BOTTOM, pady=2)
        
        # === Блок настроек нумерации ===
        settings_frame = ttk.LabelFrame(scrollable_frame, text="Настройки нумерации", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Настройка: начальная страница документа
        ttk.Label(settings_frame, text="Начать нумерацию со страницы №:", 
                 font=('Arial', 10)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.start_page_var = tk.IntVar(value=1)
        start_page_spinbox = ttk.Spinbox(
            settings_frame, 
            from_=1, 
            to=9999, 
            textvariable=self.start_page_var,
            width=10,
            font=('Arial', 10)
        )
        start_page_spinbox.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Label(settings_frame, text="(1 = первая страница)", 
                 font=('Arial', 9), foreground='gray').grid(row=0, column=2, padx=5, pady=10, sticky="w")
        
        # Настройка: начальный номер
        ttk.Label(settings_frame, text="Начать с номера:", 
                 font=('Arial', 10)).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.start_number_var = tk.IntVar(value=1)
        start_number_spinbox = ttk.Spinbox(
            settings_frame, 
            from_=1, 
            to=9999, 
            textvariable=self.start_number_var,
            width=10,
            font=('Arial', 10)
        )
        start_number_spinbox.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Label(settings_frame, text="(первая нумеруемая страница получит этот номер)", 
                 font=('Arial', 9), foreground='gray').grid(row=1, column=2, padx=5, pady=10, sticky="w")
        
        # Настройка: фон за текстом
        ttk.Label(settings_frame, text="Фон за номером:", 
                 font=('Arial', 10)).grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        self.background_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            settings_frame, 
            text="Добавлять белый фон (для лучшей читаемости)", 
            variable=self.background_var
        ).grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="w")
        
        # === Блок информации ===
        info_frame = ttk.LabelFrame(scrollable_frame, text="Информация", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.info_label = ttk.Label(
            info_frame, 
            text="Выберите входной PDF-файл для просмотра информации",
            font=('Arial', 9),
            foreground='blue'
        )
        self.info_label.pack(padx=10, pady=5)
        
        # === Кнопка запуска ===
        ttk.Button(
            scrollable_frame, 
            text="Добавить номера страниц", 
            command=self.process_pdf, 
            style='Accent.TButton'
        ).pack(pady=20)
        
        # Подсказка
        ttk.Label(
            scrollable_frame, 
            text="Можно вручную ввести пути или выбрать через диалоги",
            font=('Arial', 8), 
            foreground='gray'
        ).pack(pady=(0, 10))
        
        # Привязываем событие изменения входного файла для обновления информации
        self.input_entry.bind('<FocusOut>', self.update_file_info)
    
    def select_input_file(self):
        """Диалог выбора входного PDF-файла."""
        file_path = filedialog.askopenfilename(
            title="Выберите PDF-файл для обработки",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)
            
            # Автоматически предлагаем имя для выходного файла
            if not self.output_entry.get().strip():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_name = f"{base_name}_with_numbers.pdf"
                output_dir = os.path.dirname(file_path)
                self.output_entry.insert(0, os.path.join(output_dir, output_name))
            
            # Обновляем информацию о файле
            self.update_file_info()

    def select_output_folder(self):
        """Диалог выбора папки для сохранения."""
        folder_path = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder_path:
            # Если файл уже выбран, используем его имя, иначе предлагаем ввести
            current_output = self.output_entry.get().strip()
            if current_output:
                filename = os.path.basename(current_output)
            else:
                filename = "output.pdf"
            full_path = os.path.join(folder_path, filename)
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, full_path)

    def select_output_file(self):
        """Диалог сохранения выходного PDF-файла."""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить PDF как...",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)
    
    def update_file_info(self, event=None):
        """Обновляет информацию о выбранном PDF-файле."""
        input_file = self.input_entry.get().strip()
        
        if input_file and os.path.exists(input_file) and input_file.lower().endswith('.pdf'):
            try:
                doc = fitz.open(input_file)
                total_pages = len(doc)
                doc.close()
                
                # Обновляем максимальное значение для spinbox начальной страницы
                self.start_page_var = tk.IntVar(value=min(self.start_page_var.get(), total_pages))
                
                self.info_label.config(
                    text=f"✓ Файл загружен: {os.path.basename(input_file)} | Всего страниц: {total_pages}",
                    foreground='green'
                )
            except Exception as e:
                self.info_label.config(
                    text=f"✗ Ошибка при чтении файла: {str(e)}",
                    foreground='red'
                )
        elif input_file:
            self.info_label.config(
                text="✗ Файл не найден или не является PDF",
                foreground='red'
            )
        else:
            self.info_label.config(
                text="Выберите входной PDF-файл для просмотра информации",
                foreground='blue'
            )

    def process_pdf(self):
        """Запуск обработки PDF с проверкой введённых параметров."""
        input_file = self.input_entry.get().strip()
        output_file = self.output_entry.get().strip()

        if not input_file:
            messagebox.showwarning("Предупреждение", "Укажите входной PDF-файл!")
            return
        if not output_file:
            messagebox.showwarning("Предупреждение", "Укажите путь для сохранения результата!")
            return
        
        # Добавляем .pdf, если не указано
        if not output_file.lower().endswith('.pdf'):
            output_file += '.pdf'
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_file)
        
        # Получаем параметры нумерации
        start_page = self.start_page_var.get()
        start_number = self.start_number_var.get()
        use_background = self.background_var.get()
        
        # Проверяем, что начальная страница не превышает общее количество страниц
        try:
            doc = fitz.open(input_file)
            total_pages = len(doc)
            doc.close()
            
            if start_page > total_pages:
                messagebox.showerror(
                    "Ошибка", 
                    f"Страница начала нумерации ({start_page}) больше общего количества страниц ({total_pages})"
                )
                return
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать PDF-файл:\n{str(e)}")
            return

        try:
            add_page_numbers(input_file, output_file, start_page, start_number, use_background)
            
            # Формируем информационное сообщение о результате
            info_msg = f"✅ Номера страниц успешно добавлены!\n\n"
            info_msg += f"📄 Результат сохранён в:\n{output_file}\n\n"
            info_msg += f"📊 Параметры обработки:\n"
            info_msg += f"   • Начальная страница документа: {start_page}\n"
            info_msg += f"   • Начальный номер: {start_number}\n"
            info_msg += f"   • Фон за текстом: {'Да' if use_background else 'Нет'}"
            
            messagebox.showinfo("Успех", info_msg)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла:\n{str(e)}")


# ========== Главное окно приложения с вкладками ==========
class PDFToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Tools - Объединение и нумерация страниц")
        self.root.geometry("800x550")
        self.root.minsize(750, 500)
        
        # Создаём стиль для акцентной кнопки
        self.setup_styles()
        
        # Создаём виджет с вкладками
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаём вкладки
        self.merge_tab = MergePDFTab(self.notebook)
        self.number_tab = NumberPagesTab(self.notebook)
        
        # Добавляем вкладки в блокнот
        self.notebook.add(self.merge_tab.frame, text="Объединение PDF")
        self.notebook.add(self.number_tab.frame, text="Нумерация страниц")
        
        # Строка состояния
        self.status_bar = ttk.Label(root, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Привязка события смены вкладки
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)
    
    def setup_styles(self):
        """Настройка стилей для виджетов."""
        style = ttk.Style()
        
        # Создаём стиль для акцентной кнопки
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
        
        # Настройка цвета фона для вкладок (опционально)
        style.configure('TNotebook.Tab', padding=[10, 5], font=('Arial', 10))
    
    def on_tab_change(self, event):
        """Обновление строки состояния при смене вкладки."""
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.status_bar.config(text="Режим: объединение нескольких PDF-файлов в один")
        else:
            self.status_bar.config(text="Режим: добавление номеров страниц в PDF-документ")


# ========== Запуск приложения ==========
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolsApp(root)
    root.mainloop()

    ##pyinstaller --onefile --clean --noconfirm --hidden-import=fitz pdf_numberer.py
