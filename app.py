import os
import re
import sys
import uuid
import threading

from tkinter import (
    Tk,
    Label,
    Button,
    Listbox,
    Scrollbar,
    Frame,
    messagebox,
    filedialog,
    StringVar,
    Entry,
)
from tkinter.ttk import Progressbar, Style
import fitz

if sys.platform == "win32":
    try:
        import io

        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
    except:
        pass


def find_underscore_fields(pdf_path, exclude_rects=None):
    fields_info = []
    field_counter = 0

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        patterns = ["__", "___", "____", "_____", "______", "_______", "________"]
        found_rects = []

        for pattern in patterns:
            for rect in page.search_for(pattern):
                is_duplicate = False
                for existing in found_rects:
                    if rect.intersects(existing):
                        is_duplicate = True
                        break
                if exclude_rects:
                    for ex_rect in exclude_rects:
                        if rect.intersects(ex_rect):
                            is_duplicate = True
                            break
                if not is_duplicate:
                    found_rects.append(rect)
                    fields_info.append(
                        {
                            "page": page_num,
                            "rect": rect,
                            "field_name": f"field_{field_counter + 1}",
                        }
                    )
                    field_counter += 1

    doc.close()
    return fields_info


def find_table_fields(
    pdf_path,
    skip_first_column=False,
    skip_first_row=True,
    inset=2,
    min_cell_width=30,
    min_cell_height=15,
):
    fields_info = []
    field_counter = 0

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        finder = page.find_tables()
        edges = finder.edges

        if len(edges) < 4:
            continue

        h_edges = [e for e in edges if e["orientation"] == "h"]
        v_edges = [e for e in edges if e["orientation"] == "v"]

        if not h_edges or not v_edges:
            continue

        y_coords = sorted(set(e["top"] for e in h_edges))
        x_coords = sorted(set(e["x0"] for e in v_edges))

        if len(y_coords) < 2 or len(x_coords) < 2:
            continue

        for row_idx in range(len(y_coords) - 1):
            for col_idx in range(len(x_coords) - 1):
                if skip_first_row and row_idx == 0:
                    continue
                if skip_first_column and col_idx == 0:
                    continue

                rect = fitz.Rect(
                    x_coords[col_idx] + inset,
                    y_coords[row_idx] + inset,
                    x_coords[col_idx + 1] - inset,
                    y_coords[row_idx + 1] - inset,
                )

                if rect.width >= min_cell_width and rect.height >= min_cell_height:
                    fields_info.append(
                        {
                            "page": page_num,
                            "rect": rect,
                            "field_name": f"table_{field_counter + 1}",
                        }
                    )
                    field_counter += 1

    doc.close()
    return fields_info


def add_pdf_fields(pdf_path, fields_info, output_path):
    doc = fitz.open(pdf_path)

    for field in fields_info:
        page_num = field["page"]
        if page_num >= len(doc):
            continue

        page = doc.load_page(page_num)
        rect = field["rect"]

        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_name = field["field_name"]
        widget.rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1)
        widget.fill_color = None
        widget.border_width = 0
        widget.text_color = fitz.pdfcolor["black"]
        widget.text_font = "Helv"
        widget.text_fontsize = max(8, (rect.height - 4))

        page.add_widget(widget)

    doc.save(output_path)
    doc.close()


def convert_to_pdf(input_docx, output_pdf):
    try:
        from docx2pdf import convert

        convert(input_docx, output_pdf)
        return os.path.exists(output_pdf)
    except ImportError:
        messagebox.showerror("Ошибка", "Установите docx2pdf: pip install docx2pdf")
        return False
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка конвертации: {e}")
        return False


def process_file(input_file, output_folder, progress_callback=None, log_errors=None):
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    temp_pdf = os.path.join(output_folder, f"temp_{uuid.uuid4()}.pdf")
    final_pdf = os.path.join(output_folder, f"{base_name}.pdf")

    try:
        if progress_callback:
            progress_callback(f"Конвертация в PDF...")

        if not convert_to_pdf(input_file, temp_pdf):
            err = "Ошибка конвертации DOCX в PDF"
            if log_errors:
                log_errors(base_name, err)
            return {"success": False, "error": err, "file": base_name}

        if progress_callback:
            progress_callback(f"Поиск полей...")

        underscore_fields = find_underscore_fields(temp_pdf)
        table_fields = find_table_fields(temp_pdf)

        fields_info = underscore_fields + table_fields

        if not fields_info:
            err = f"Подчеркивания (___) и таблицы не найдены"
            if log_errors:
                log_errors(base_name, err)
            return {"success": False, "error": err, "file": base_name}

        if progress_callback:
            progress_callback(f"Добавление полей...")

        add_pdf_fields(temp_pdf, fields_info, final_pdf)

        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)

        return {"success": True, "file": final_pdf, "count": len(fields_info)}

    except Exception as e:
        if log_errors:
            log_errors(base_name, str(e))
        return {"success": False, "error": str(e), "file": base_name}


class App:
    def __init__(self):
        self.root = Tk()
        self.root.title("Docx -> PDF s polyami")
        self.root.geometry("1300x820")
        self.root.configure(bg="#f0f4f8")

        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(2)
        except:
            pass

        try:
            self.root.tk.call("tk", "scaling", 1.5)
        except:
            pass

        Style().theme_use("clam")
        self.files = []
        self.output_folder = StringVar()
        self.setup_ui()

    def setup_ui(self):
        main = Frame(self.root, bg="white")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        Label(
            main,
            text="Docx - PDF s polyami",
            font=("Segoe UI", 24, "bold"),
            bg="white",
            fg="#1e3a5f",
        ).pack(pady=(10, 5))
        Label(
            main,
            text="Конвертирует подчеркивания (___) в текстовые поля",
            font=("Segoe UI", 11),
            bg="white",
            fg="#64748b",
        ).pack(pady=(0, 20))

        Label(
            main,
            text="Файлы для конвертации:",
            font=("Segoe UI", 13, "bold"),
            bg="white",
            fg="#1e293b",
        ).pack(anchor="w", pady=(0, 8))

        list_border = Frame(main, bg="#e2e8f0", bd=2, relief="solid")
        list_border.pack(fill="both", expand=True)
        list_inner = Frame(list_border, bg="white")
        list_inner.pack(fill="both", expand=True, padx=2, pady=2)

        scrollbar = Scrollbar(list_inner, orient="vertical")
        self.file_list = Listbox(
            list_inner,
            font=("Segoe UI", 11),
            bg="#f8fafc",
            fg="#1e293b",
            yscrollcommand=scrollbar.set,
            height=8,
            bd=0,
            highlightthickness=0,
        )
        scrollbar.config(command=self.file_list.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 5))
        self.file_list.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        btn_frame = Frame(main, bg="white")
        btn_frame.pack(fill="x", pady=(10, 0))
        Button(
            btn_frame,
            text="Добавить файлы",
            font=("Segoe UI", 11, "bold"),
            command=self.add_files,
            bg="#3b82f6",
            fg="white",
            padx=20,
            pady=8,
            bd=0,
            relief="flat",
            cursor="hand2",
        ).pack(side="left", padx=(0, 10))
        Button(
            btn_frame,
            text="Удалить выбранный",
            font=("Segoe UI", 11),
            command=self.remove_file,
            bg="#ef4444",
            fg="white",
            padx=15,
            pady=8,
            bd=0,
            relief="flat",
            cursor="hand2",
        ).pack(side="left")

        Label(
            main,
            text="Папка для сохранения:",
            font=("Segoe UI", 13, "bold"),
            bg="white",
            fg="#1e293b",
        ).pack(anchor="w", pady=(15, 8))
        folder_select = Frame(main, bg="white")
        folder_select.pack(fill="x")
        Entry(
            folder_select,
            textvariable=self.output_folder,
            font=("Segoe UI", 11),
            bg="#f1f5f9",
            fg="#1e293b",
            bd=2,
            relief="solid",
            state="readonly",
            readonlybackground="#f1f5f9",
        ).pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=8)
        Button(
            folder_select,
            text="Выбрать папку",
            font=("Segoe UI", 11, "bold"),
            command=self.select_folder,
            bg="#10b981",
            fg="white",
            padx=20,
            pady=8,
            bd=0,
            relief="flat",
            cursor="hand2",
        ).pack(side="left")

        self.convert_btn = Button(
            main,
            text="Конвертировать в PDF",
            font=("Segoe UI", 14, "bold"),
            command=self.start_conversion,
            bg="#1e3a5f",
            fg="white",
            padx=40,
            pady=12,
            bd=0,
            relief="flat",
            cursor="hand2",
        )
        self.convert_btn.pack(pady=15)

        self.progress_frame = Frame(main, bg="white")
        self.progress_bar = Progressbar(
            self.progress_frame, mode="indeterminate", length=300
        )
        self.status_label = Label(
            main, text="", font=("Segoe UI", 10), bg="white", fg="#64748b"
        )

        info = "Инструкция:\n- Добавьте файлы .docx с подчеркиваниями (__)\n- Выберите папку для сохранения\n- Нажмите Конвертировать"
        Label(
            main,
            text=info,
            font=("Segoe UI", 10),
            bg="white",
            fg="#64748b",
            justify="left",
            padx=15,
            pady=10,
        ).pack(pady=(10, 0))

    def add_files(self):
        filenames = filedialog.askopenfilenames(
            title="Выберите файлы DOCX", filetypes=[("DOCX files", "*.docx")]
        )
        for f in filenames:
            if f not in self.files:
                self.files.append(f)
                self.file_list.insert("end", os.path.basename(f))

    def remove_file(self):
        selection = self.file_list.curselection()
        if selection:
            idx = selection[0]
            self.file_list.delete(idx)
            self.files.pop(idx)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения PDF")
        if folder:
            self.output_folder.set(folder)

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("Ошибка", "Выберите файлы для конвертации")
            return
        if not self.output_folder.get():
            messagebox.showwarning("Ошибка", "Выберите папку для сохранения")
            return

        self.convert_btn.config(state="disabled")
        self.progress_frame.pack(fill="x")
        self.progress_bar.pack(fill="x")
        self.progress_bar.start()
        threading.Thread(target=self.process_files).start()

    def process_files(self):
        output_folder = self.output_folder.get()
        success_count = 0
        error_count = 0
        errors_list = []

        def log_error(filename, error):
            errors_list.append(f"{filename}: {error}")

        def update(msg):
            self.root.after(0, lambda m=msg: self.status_label.config(text=m))
            self.root.after(0, lambda: self.status_label.pack(pady=5))

        for i, f in enumerate(self.files):
            update(f"Обработка {i + 1}/{len(self.files)}: {os.path.basename(f)}")
            result = process_file(f, output_folder, update, log_error)
            if result["success"]:
                success_count += 1
            else:
                error_count += 1

        self.root.after(0, self.progress_bar.stop)
        self.root.after(0, self.progress_frame.pack_forget)
        self.root.after(0, lambda: self.status_label.pack_forget())
        self.root.after(0, lambda: self.convert_btn.config(state="normal"))

        if errors_list:
            msg = "Ошибки:\n" + "\n".join(errors_list[:5])
            self.root.after(0, lambda: messagebox.showerror("Ошибки", msg))
        else:
            msg = f"Готово! Успешно: {success_count}, Ошибок: {error_count}"
            self.root.after(0, lambda: messagebox.showinfo("Результат", msg))

        self.root.after(0, lambda: self.status_label.config(text=""))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
