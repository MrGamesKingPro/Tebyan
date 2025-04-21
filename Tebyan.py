# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, font as tkfont
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import arabic_reshaper # Import the library
from bidi.algorithm import get_display
import sys
import os
import json
import csv
import xml.etree.ElementTree as ET # <--- Import ElementTree
from xml.etree.ElementTree import ParseError as ETParseError # <-- Import specific ParseError
import threading
import queue # For thread-safe GUI updates

# --- Constants and Arabic UI Strings ---
APP_TITLE = "برنامج تبيان"
INPUT_LABEL_TEXT = ": أدخل النص الأصلي العربي هنا"
OUTPUT_LABEL_TEXT = ": النص المعدل (انسخ من هنا)"
PROCESS_BUTTON_TEXT = "معالجة النص"
COPY_ALL_BUTTON_TEXT = "نسخ الكل"
CLEAR_ALL_TEXT_BUTTON_TEXT = "مسح الكل (النصوص)"
DIRECT_TAB_TEXT = "معالجة نصوص"
FILE_TAB_TEXT = "معالجة ملفات"
SELECT_FILES_BUTTON_TEXT = "فتح ملفًا أو أكثر"
SELECTED_FILES_LABEL_TEXT = ":عدد الملفات المحددة"
SELECT_OUTPUT_DIR_BUTTON_TEXT = "مسار المجلد الحفظ"
SELECTED_OUTPUT_DIR_LABEL_TEXT = ": مجلد الحفظ"
PROCESS_FILES_BUTTON_TEXT = "معالجة الملفات المحددة"
CLEAR_ALL_FILES_BUTTON_TEXT = "مسح الكل (الملفات)"
STATUS_LABEL_TEXT = "ملفات :"
READY_STATUS = "سجل"
PROCESSING_STATUS = "جاري المعالجة..."
DONE_STATUS = "اكتملت المعالجة."
ERROR_STATUS = "حدث خطأ."
MENU_FILE = "ملف"
MENU_SAVE_TEXT = "حفظ النص المعالج..."
MENU_EXIT = "خروج"
MENU_HELP = "مساعدة"
MENU_ABOUT = "عن البرنامج"
MENU_INSTRUCTIONS = "تعليمات"
HELP_TITLE = "تعليمات"
ABOUT_TITLE = "حول البرنامج"
ABOUT_TEXT = f"{APP_TITLE}\n\nالإصدار: 1.2\n\nبرنامج لمعالجة النصوص العربية لعرضها بشكل صحيح في التطبيقات والألعاب التي لا تدعم اللغة العربية.\n\n MrGamesKingPro Ⓒ 2025  جميع الحقوق محفوظة \n\n  https://github.com/MrGamesKingPro" # Version Bumped to 1.2
HELP_TEXT = """
كيفية الاستخدام:

1.  **معالجة نصوص**
    *   .اكتب أو الصق النص العربي في المربع العلوي.
    *   .اضغط على زر 'معالجة النص'
    *   .سيظهر النص المعدل في المربع السفلي، جاهزاً للنسخ.
    *   .يمكنك نسخ النص المعدل بالكامل باستخدام زر 'نسخ الكل' الموجود أسفله.
    *   .يمكنك مسح مربعات النص الإدخال والإخراج باستخدام زر 'مسح الكل (النصوص)'
    *   .يمكنك حفظ النص المعدل من قائمة 'ملف'

2.  **معالجة ملفات**
    *   .(.txt)، JSON (.json)، CSV (.csv)، أو XML (.xml) اضغط على 'فتح ملفًا أو أكثر' لتحديد ملفات نصية ذات إمتداد
    *   .اضغط على 'مسارالمجلد الحفظ' لتحديد المجلد الذي سيتم حفظ الملفات المعالجة فيه
    *   .سيتم إنشاء مجلد فرعي باسم 'processed_files' داخل المجلد المختار (إذا لم يكن موجوداً)
    *   .اضغط على 'معالجة الملفات المحددة'.
    *   .سيتم معالجة النصوص العربية فقط داخل الملفات وحفظ نسخ جديدة بنفس إسم والتنسيق في المجلد المخصص
    *   .(XML سيتم معالجة النصوص داخل العناصر، ونهايات العناصر، وقيم السمات في ملفات)
    *   .تابع شريط التقدم وسجل الحالة لمعرفة حالة العملية.
    *   .يمكنك مسح قائمة الملفات ومجلد الحفظ والسجل باستخدام زر 'مسح الكل (الملفات)'

ملاحظات:
*    .UTF-8 تأكد من أن الملفات تستخدم ترميز
*   .في ملفات CSV و JSON، سيتم محاولة معالجة كل القيم النصية
*   .في ملفات XML، سيتم محاولة معالجة كل النصوص الموجودة في العناصر والسمات ونهايات العناصر
*   .يمكنك استخدام زر الفأرة الأيمن للقص/النسخ/اللصق في مربعات النص (حسب ما إذا كان المربع قابلاً للكتابة أو للقراءة فقط)
"""
COPY_TEXT = "نسخ"
PASTE_TEXT = "لصق"
CUT_TEXT = "قص"
SELECT_ALL_TEXT = "تحديد الكل"
ERROR_READING_FILE = "خطأ في قراءة الملف"
ERROR_WRITING_FILE = "خطأ في كتابة الملف"
ERROR_PROCESSING_FILE = "خطأ في معالجة الملف"
ERROR_INVALID_FORMAT = "تنسيق ملف غير صالح أو غير مدعوم"
ERROR_SELECT_FILES = "الرجاء اختيار ملف واحد على الأقل."
ERROR_SELECT_OUTPUT_DIR = "الرجاء اختيار مجلد للحفظ."
FILE_PROCESSED_SUCCESS = "تمت معالجة الملف بنجاح:"
OUTPUT_FOLDER_NAME = "processed_files" # Name for the subfolder

# --- Core Processing Function ---
# ... (Keep process_arabic_text as is) ...
def process_arabic_text(input_text):
    if not isinstance(input_text, str) or not input_text:
        return input_text # Return non-strings or empty strings as is
    try:
        has_arabic = any('\u0600' <= char <= '\u06FF' or # Arabic
                         '\u0750' <= char <= '\u077F' or # Arabic Supplement
                         '\u08A0' <= char <= '\u08FF' or # Arabic Extended-A
                         '\uFB50' <= char <= '\uFDFF' or # Arabic Presentation Forms-A
                         '\uFE70' <= char <= '\uFEFF'    # Arabic Presentation Forms-B
                         for char in input_text)
        if not has_arabic:
            return input_text # Return text without Arabic unchanged
        configuration = {
            'delete_harakat': False,
            'shift_harakat_position': False,
            'support_ligatures': True,
        }
        reshaper_instance = arabic_reshaper.ArabicReshaper(configuration=configuration)
        reshaped_text = reshaper_instance.reshape(input_text)
        bidi_text = get_display(reshaped_text)
        return bidi_text
    except Exception as e:
        print(f"Error processing text segment: '{input_text[:50]}...' - {e}", file=sys.stderr)
        return input_text # Return original on error during processing segment

# --- File Processing Logic ---
# ... (Keep process_txt_file, process_csv_file, _process_json_value, process_json_file as is) ...
def process_txt_file(input_path, output_path):
    try:
        with open(input_path, 'r', encoding='utf-8') as infile, \
             open(output_path, 'w', encoding='utf-8') as outfile:
            for line in infile:
                processed_line = process_arabic_text(line.rstrip('\r\n'))
                outfile.write(processed_line + '\n')
        return True, None
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}: {e}"

def process_csv_file(input_path, output_path):
    try:
        rows = []
        with open(input_path, 'r', encoding='utf-8', newline='') as infile:
            reader = csv.reader(infile)
            try:
                header = next(reader)
                rows.append([process_arabic_text(cell) if isinstance(cell, str) else cell for cell in header])
            except StopIteration:
                 return True, None # Empty file is okay
            for row in reader:
                processed_row = [process_arabic_text(cell) if isinstance(cell, str) else cell for cell in row]
                rows.append(processed_row)
        with open(output_path, 'w', encoding='utf-8', newline='') as outfile:
            writer = csv.writer(outfile)
            writer.writerows(rows)
        return True, None
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (CSV): {e}"

def _process_json_value(value):
    if isinstance(value, str):
        return process_arabic_text(value)
    elif isinstance(value, dict):
        return {k: _process_json_value(v) for k, v in value.items()}
    elif isinstance(value, list):
        return [_process_json_value(item) for item in value]
    else:
        return value

def process_json_file(input_path, output_path):
    try:
        with open(input_path, 'r', encoding='utf-8') as infile:
            data = json.load(infile)
        processed_data = _process_json_value(data)
        with open(output_path, 'w', encoding='utf-8') as outfile:
            json.dump(processed_data, outfile, ensure_ascii=False, indent=4)
        return True, None
    except json.JSONDecodeError as e:
         return False, f"{ERROR_READING_FILE} (JSON Decode): {e}"
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (JSON): {e}"

# --- NEW: XML File Processing Logic ---
def process_xml_file(input_path, output_path):
    """
    Parses an XML file, processes Arabic text in elements and attributes,
    and writes the modified XML to the output path.
    """
    try:
        # Parse the XML file
        # We need a parser that recovers from errors to handle comments/PIs correctly sometimes
        parser = ET.XMLParser(encoding='utf-8', target=ET.TreeBuilder())
        tree = ET.parse(input_path, parser=parser)
        root = tree.getroot()

        # Iterate through ALL elements in the tree using iter()
        for element in root.iter():
            # 1. Process element.text (text directly within tags)
            if element.text:
                element.text = process_arabic_text(element.text)

            # 2. Process element.tail (text after the element's closing tag)
            if element.tail:
                element.tail = process_arabic_text(element.tail)

            # 3. Process attribute values
            for attr_key, attr_value in element.attrib.items():
                if isinstance(attr_value, str): # Check if value is string
                    element.attrib[attr_key] = process_arabic_text(attr_value)

        # Write the modified tree back to a file
        # xml_declaration=True adds <?xml version="1.0" encoding="UTF-8"?>
        # method="xml" ensures standard XML output (not HTML, etc.)
        tree.write(output_path, encoding='utf-8', xml_declaration=True, method="xml")
        return True, None
    except ETParseError as e:
        return False, f"{ERROR_READING_FILE} (XML Parse): {e}"
    except Exception as e:
        # Catch other potential errors during processing or writing
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (XML): {e}"

# --- Worker Thread for File Processing ---
# ... (Updated file_processing_worker) ...
def file_processing_worker(file_list, output_dir, progress_queue):
    total_files = len(file_list)
    processed_count = 0
    processed_output_dir = os.path.join(output_dir, OUTPUT_FOLDER_NAME)
    try:
        os.makedirs(processed_output_dir, exist_ok=True)
    except OSError as e:
        progress_queue.put(("error", f"خطأ في إنشاء مجلد الحفظ: {processed_output_dir} - {e}"))
        progress_queue.put(("done", None))
        return

    for input_path in file_list:
        if not progress_queue.full(): # Rough check for GUI responsiveness
            pass
        else:
            print("GUI queue full or unresponsive, stopping file processing thread early.", file=sys.stderr)
            progress_queue.put(("error", "تم إيقاف المعالجة مبكرًا بسبب عدم استجابة الواجهة."))
            progress_queue.put(("done", None))
            return

        try:
            base_name = os.path.basename(input_path)
            name, ext = os.path.splitext(base_name)
            safe_name = name.replace(" ", "_")
            output_filename = f"{safe_name}_processed{ext}"
            output_path = os.path.join(processed_output_dir, output_filename)
            success = False
            error_msg = None
            ext_lower = ext.lower()
            progress_queue.put(("log", f"معالجة: {base_name} ...")) # Log start
            if ext_lower == '.txt':
                success, error_msg = process_txt_file(input_path, output_path)
            elif ext_lower == '.csv':
                 success, error_msg = process_csv_file(input_path, output_path)
            elif ext_lower == '.json':
                 success, error_msg = process_json_file(input_path, output_path)
            elif ext_lower == '.xml': # <--- Added XML check
                 success, error_msg = process_xml_file(input_path, output_path)
            else:
                 error_msg = f"{ERROR_INVALID_FORMAT}: {ext}"
                 success = False
            processed_count += 1
            progress = int((processed_count / total_files) * 100) if total_files > 0 else 100
            if success:
                 progress_queue.put(("log", f"{FILE_PROCESSED_SUCCESS} {output_filename}"))
                 progress_queue.put(("progress", progress))
            else:
                 error_detail = error_msg if error_msg else "Unknown error"
                 progress_queue.put(("error", f"{ERROR_PROCESSING_FILE} '{base_name}': {error_detail}"))
                 progress_queue.put(("progress", progress)) # Update progress even on error
        except Exception as e:
            processed_count += 1
            progress = int((processed_count / total_files) * 100) if total_files > 0 else 100
            progress_queue.put(("error", f"Unexpected error processing '{os.path.basename(input_path)}': {e}"))
            progress_queue.put(("progress", progress))

    progress_queue.put(("done", None))


# --- GUI Application Class ---
class ArabicProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.style = tb.Style(theme="litera")
        self.root.minsize(650, 630)

        # Font setup (same as before)
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.arabic_font_family = "Tahoma"
        self.default_font_size = self.default_font.actual()["size"]
        try:
             tkfont.Font(family=self.arabic_font_family)
        except tk.TclError:
             print(f"Font '{self.arabic_font_family}' not found, falling back to system default '{self.default_font.actual()['family']}'", file=sys.stderr)
             self.arabic_font_family = self.default_font.actual()["family"]
        self.arabic_ui_font = tkfont.Font(family=self.arabic_font_family, size=self.default_font_size + 1)
        self.arabic_text_font = tkfont.Font(family=self.arabic_font_family, size=self.default_font_size + 2)
        self.style.configure('.', font=self.arabic_ui_font)
        self.menu_font_spec = (self.arabic_font_family, self.arabic_ui_font.actual()["size"])
        self.listbox_font_spec = (self.arabic_font_family, self.arabic_ui_font.actual()["size"])
        self.text_widget_font_spec = (self.arabic_font_family, self.arabic_text_font.actual()["size"])

        self.selected_files = []
        self.output_directory = ""
        self.progress_queue = queue.Queue()
        self.worker_thread = None

        # --- Menu Bar --- (same as before)
        self.menu_bar = tk.Menu(root)
        root.config(menu=self.menu_bar)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.menu_font_spec)
        self.menu_bar.add_cascade(label=MENU_FILE, menu=self.file_menu)
        self.file_menu.add_command(label=MENU_SAVE_TEXT, command=self.save_processed_text)
        self.file_menu.add_separator()
        self.file_menu.add_command(label=MENU_EXIT, command=self.on_closing)
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0, font=self.menu_font_spec)
        self.menu_bar.add_cascade(label=MENU_HELP, menu=self.help_menu)
        self.help_menu.add_command(label=MENU_INSTRUCTIONS, command=self.show_help)
        self.help_menu.add_command(label=MENU_ABOUT, command=self.show_about)

        # --- Main Layout (Notebook for Tabs) ---
        self.notebook = ttk.Notebook(root, bootstyle="primary")
        self.notebook.pack(pady=10, padx=10, expand=True, fill=BOTH)

        # --- Tab 1: Direct Text Processing ---
        # ... (Tab 1 remains unchanged) ...
        self.direct_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.direct_tab, text=DIRECT_TAB_TEXT)
        input_label = ttk.Label(self.direct_tab, text=INPUT_LABEL_TEXT)
        input_label.pack(pady=(0, 5), anchor=E)
        self.input_text_area = scrolledtext.ScrolledText(self.direct_tab, height=8, width=60, wrap=tk.WORD, font=self.text_widget_font_spec)
        self.input_text_area.pack(pady=5, padx=5, expand=True, fill=BOTH)
        self.input_text_area.configure(bd=1, relief="sunken")
        self.input_text_area.config(insertbackground='black')
        self.input_text_area.tag_configure("right", justify='right')
        self.input_text_area.tag_add("right", "1.0", "end")
        self.input_text_area.bind("<KeyRelease>", self._force_right_align_input)
        self.input_text_area.bind("<Return>", self.handle_input_enter)
        self.input_text_area.focus()
        self.add_context_menu(self.input_text_area)
        process_button = tb.Button(self.direct_tab, text=PROCESS_BUTTON_TEXT, command=self.process_direct_text, bootstyle="success")
        process_button.pack(pady=10)
        output_label = ttk.Label(self.direct_tab, text=OUTPUT_LABEL_TEXT)
        output_label.pack(pady=(5, 5), anchor=E)
        self.output_text_area = scrolledtext.ScrolledText(self.direct_tab, height=8, width=60, wrap=tk.WORD, font=self.text_widget_font_spec)
        self.output_text_area.configure(bd=1, relief="sunken")
        self.output_text_area.tag_configure("right_readonly", justify='right')
        self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
        self.add_context_menu(self.output_text_area)
        self.output_text_area.pack(pady=5, padx=5, expand=True, fill=BOTH)
        direct_button_frame = ttk.Frame(self.direct_tab)
        direct_button_frame.pack(pady=(5, 0), fill=X)
        self.clear_direct_button = tb.Button(direct_button_frame, text=CLEAR_ALL_TEXT_BUTTON_TEXT, command=self.clear_direct_tab, bootstyle="warning")
        self.clear_direct_button.pack(side=LEFT, padx=(0, 5))
        self.copy_all_button = tb.Button(direct_button_frame, text=COPY_ALL_BUTTON_TEXT, command=self.copy_all_output, bootstyle="secondary", state=tk.DISABLED)
        self.copy_all_button.pack(side=RIGHT, padx=(5, 0))


        # --- Tab 2: File Processing ---
        self.file_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.file_tab, text=FILE_TAB_TEXT)

        # File Selection Frame (same as before)
        file_select_frame = ttk.Frame(self.file_tab)
        file_select_frame.pack(fill=X, pady=5)
        select_files_button = tb.Button(file_select_frame, text=SELECT_FILES_BUTTON_TEXT, command=self.select_files, bootstyle="info")
        select_files_button.pack(side=RIGHT, padx=(10, 0))
        self.selected_files_label = ttk.Label(file_select_frame, text=f"({len(self.selected_files)}) {SELECTED_FILES_LABEL_TEXT}")
        self.selected_files_label.pack(side=RIGHT, fill=X, expand=True, padx=(0, 5))

        # Listbox Frame for Selected Files (same as before)
        self.files_listbox_frame = ttk.Frame(self.file_tab)
        self.files_listbox_frame.pack(pady=5, padx=5, fill=X, side=TOP)
        self.files_listbox_scrollbar = ttk.Scrollbar(self.files_listbox_frame, orient=VERTICAL)
        self.files_listbox_scrollbar.pack(side=RIGHT, fill=Y)
        self.files_listbox = tk.Listbox(self.files_listbox_frame, height=5, width=70,
                                        yscrollcommand=self.files_listbox_scrollbar.set,
                                        font=self.listbox_font_spec,
                                        justify=tk.RIGHT,
                                        exportselection=False)
        self.files_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.files_listbox_scrollbar.config(command=self.files_listbox.yview)
        self.add_context_menu(self.files_listbox)
        self.update_selected_files_display()

        # Output Directory Frame (same as before)
        output_dir_frame = ttk.Frame(self.file_tab)
        output_dir_frame.pack(fill=X, pady=5)
        select_output_dir_button = tb.Button(output_dir_frame, text=SELECT_OUTPUT_DIR_BUTTON_TEXT, command=self.select_output_dir, bootstyle="info")
        select_output_dir_button.pack(side=RIGHT, padx=(10, 0))
        self.selected_output_dir_label = ttk.Label(output_dir_frame, text=f"- :{SELECTED_OUTPUT_DIR_LABEL_TEXT}", wraplength=450, anchor=E, justify=RIGHT)
        self.selected_output_dir_label.pack(side=RIGHT, fill=X, expand=True, padx=(0, 5))

        # --- Button Frame for Process Files and Clear All (Files Tab) ---
        file_action_button_frame = ttk.Frame(self.file_tab)
        file_action_button_frame.pack(pady=10, fill=X)

        # Clear All (Files) Button - Pack LEFT
        self.clear_files_button = tb.Button(file_action_button_frame, text=CLEAR_ALL_FILES_BUTTON_TEXT, command=self.clear_file_tab, bootstyle="warning")
        self.clear_files_button.pack(side=LEFT, padx=(0, 5)) # Align left, add padding right

        # Process Files Button - Pack RIGHT
        self.process_files_button = tb.Button(file_action_button_frame, text=PROCESS_FILES_BUTTON_TEXT, command=self.start_file_processing, bootstyle="success")
        self.process_files_button.pack(side=RIGHT, padx=(5, 0)) # Align right, add padding left

        # Status Area Frame (same as before)
        status_frame = ttk.Frame(self.file_tab)
        status_frame.pack(fill=X, pady=(10, 0))
        self.status_label = ttk.Label(status_frame, text=READY_STATUS, anchor=E)
        self.status_label.pack(side=RIGHT, padx=(5, 0))
        status_label_title = ttk.Label(status_frame, text=STATUS_LABEL_TEXT, anchor=E)
        status_label_title.pack(side=RIGHT)

        # Progress Bar (same as before)
        self.progress_bar = tb.Progressbar(self.file_tab, mode='determinate', bootstyle="success-striped")
        self.progress_bar.pack(pady=5, padx=5, fill=X)

        # Log Area (same as before)
        log_font_spec = (self.arabic_font_family, self.arabic_ui_font.actual()["size"] - 1)
        self.log_area = scrolledtext.ScrolledText(self.file_tab, height=6, width=70, wrap=tk.WORD, state=tk.DISABLED, font=log_font_spec)
        self.log_area.pack(pady=5, padx=5, fill=BOTH, expand=True)
        self.log_area.tag_configure("error", foreground="red")
        self.log_area.tag_configure("success", foreground="#008000")
        self.log_area.tag_configure("info", foreground="blue")
        self.log_area.tag_configure("log_base", justify='right')
        self.log_area.config(cursor="arrow")
        self.add_context_menu(self.log_area)

        # Start checking the progress queue
        self.check_progress_queue()

    # --- Force Right Alignment for Input Area --- (same as before)
    def _force_right_align_input(self, event=None):
        try:
            if not self.input_text_area.winfo_exists(): return
            self.input_text_area.tag_remove("right", "1.0", "end")
            self.input_text_area.tag_add("right", "1.0", "end")
        except tk.TclError:
            pass

    # --- Handler for Enter key in Input Area --- (same as before)
    def handle_input_enter(self, event=None):
        try:
            if not self.input_text_area.winfo_exists(): return 'break'
            self.input_text_area.mark_set(tk.INSERT, tk.END)
            self.input_text_area.insert(tk.INSERT, '\n')
            self.input_text_area.see(tk.END)
            self._force_right_align_input()
            self.input_text_area.focus_set()
            return 'break'
        except tk.TclError:
            return 'break'
        except Exception as e:
             print(f"Error in handle_input_enter: {e}", file=sys.stderr)
             return 'break'

    # --- GUI Event Handlers ---
    # ... (process_direct_text, clear_direct_tab, copy_all_output, save_processed_text remain unchanged) ...
    def process_direct_text(self):
        """Handles processing text from the direct input area."""
        has_processed_text = False
        try:
            if not self.input_text_area.winfo_exists() or not self.output_text_area.winfo_exists():
                return
            input_text = self.input_text_area.get("1.0", tk.END).strip()

            self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
            self.output_text_area.delete("1.0", tk.END)
            self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
            if self.copy_all_button.winfo_exists():
                 self.copy_all_button.config(state=tk.DISABLED)

            if not input_text:
                self._force_right_align_input()
                if self.direct_tab.winfo_exists():
                   messagebox.showwarning(APP_TITLE, "الرجاء إدخال نص للمعالجة.", parent=self.direct_tab)
                return

            processed_text = process_arabic_text(input_text)
            self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
            self.output_text_area.insert(tk.END, processed_text)
            self.output_text_area.tag_remove("right_readonly", "1.0", "end")
            self.output_text_area.tag_add("right_readonly", "1.0", "end")
            self.output_text_area.config(state=tk.DISABLED, cursor="arrow")

            has_processed_text = bool(processed_text)
            if has_processed_text and self.copy_all_button.winfo_exists():
                self.copy_all_button.config(state=tk.NORMAL)

        except tk.TclError:
             try:
                 if self.direct_tab.winfo_exists():
                     messagebox.showerror(APP_TITLE, "حدث خطأ أثناء الوصول لواجهة المستخدم.", parent=self.direct_tab)
             except tk.TclError: pass
        except Exception as e:
            try:
                if self.direct_tab.winfo_exists():
                    messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع:\n{e}", parent=self.direct_tab)
            except tk.TclError: pass
        finally:
            try:
                if not has_processed_text and self.copy_all_button.winfo_exists():
                    self.copy_all_button.config(state=tk.DISABLED)
                if self.output_text_area.winfo_exists() and self.output_text_area.cget('state') == tk.NORMAL:
                     self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
            except tk.TclError:
                pass

    def clear_direct_tab(self):
        """Clears the input and output text areas in the direct tab."""
        try:
            if self.input_text_area.winfo_exists():
                self.input_text_area.delete("1.0", tk.END)
                self._force_right_align_input()
            if self.output_text_area.winfo_exists():
                self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
                self.output_text_area.delete("1.0", tk.END)
                self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
            if self.copy_all_button.winfo_exists():
                self.copy_all_button.config(state=tk.DISABLED)
        except tk.TclError:
             try:
                 if self.direct_tab.winfo_exists():
                     messagebox.showerror(APP_TITLE, "حدث خطأ أثناء مسح مربعات النص.", parent=self.direct_tab)
             except tk.TclError: pass
        except Exception as e:
             try:
                 if self.direct_tab.winfo_exists():
                     messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع أثناء المسح:\n{e}", parent=self.direct_tab)
             except tk.TclError: pass

    def copy_all_output(self):
        text_to_copy = ""
        try:
            if not self.output_text_area.winfo_exists() or not self.root.winfo_exists():
                return
            self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
            text_to_copy = self.output_text_area.get("1.0", tk.END).strip()
        except tk.TclError as e:
            print(f"TclError getting text from output area: {e}", file=sys.stderr)
        finally:
            try:
                if self.output_text_area.winfo_exists():
                    self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
            except tk.TclError:
                 pass
        if text_to_copy and self.root.winfo_exists():
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(text_to_copy)
                self.root.update()
            except tk.TclError as e:
                print(f"TclError interacting with clipboard: {e}", file=sys.stderr)
                try:
                    if self.direct_tab.winfo_exists():
                         messagebox.showerror(APP_TITLE, f"خطأ في النسخ إلى الحافظة:\n{e}", parent=self.direct_tab)
                except tk.TclError: pass

    def save_processed_text(self):
        processed_text = ""
        try:
             if not self.output_text_area.winfo_exists(): return
             self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
             processed_text = self.output_text_area.get("1.0", tk.END).strip()
             self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
             if not processed_text:
                 if self.root.winfo_exists():
                     messagebox.showwarning(APP_TITLE, "لا يوجد نص معالج للحفظ.", parent=self.root)
                 return
             parent_widget = self.root if self.root.winfo_exists() else None
             if not parent_widget: return
             file_path = filedialog.asksaveasfilename(
                 defaultextension=".txt",
                 filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
                 title="حفظ النص المعالج",
                 parent=parent_widget
             )
             if file_path:
                 try:
                     processed_text_save = processed_text.replace('\r\n', '\n').replace('\r', '\n')
                     with open(file_path, 'w', encoding='utf-8') as f:
                         f.write(processed_text_save + '\n')
                     messagebox.showinfo(APP_TITLE, f"تم حفظ النص بنجاح في:\n{file_path}", parent=parent_widget)
                 except Exception as e:
                     messagebox.showerror(APP_TITLE, f"{ERROR_WRITING_FILE}:\n{e}", parent=parent_widget)
        except tk.TclError:
            try:
                if self.root.winfo_exists():
                   messagebox.showerror(APP_TITLE, "حدث خطأ أثناء الوصول لواجهة المستخدم.", parent=self.root)
                if self.output_text_area.winfo_exists() and self.output_text_area.cget('state') == tk.NORMAL:
                    self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
            except tk.TclError: pass
        except Exception as e:
             try:
                 if self.root.winfo_exists():
                    messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع:\n{e}", parent=self.root)
                 if self.output_text_area.winfo_exists() and self.output_text_area.cget('state') == tk.NORMAL:
                    self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
             except tk.TclError: pass

    # --- Select Files --- (Updated for XML)
    def select_files(self):
        parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
        if not parent_widget.winfo_exists(): return
        try:
            selected = filedialog.askopenfilenames(
                title="اختر ملفات للمعالجة",
                filetypes=[("Text Files", "*.txt"),
                           ("CSV Files", "*.csv"),
                           ("JSON Files", "*.json"),
                           ("XML Files", "*.xml"), # <--- Added XML type
                           ("All Files", "*.*")],
                 parent=parent_widget
            )
            if selected:
                self.selected_files = list(selected)
            self.update_selected_files_display()
        except Exception as e:
             print(f"Error during file selection: {e}", file=sys.stderr)
             try:
                messagebox.showerror(APP_TITLE, f"حدث خطأ أثناء اختيار الملفات:\n{e}", parent=parent_widget)
             except tk.TclError: pass

    # --- Update Selected Files Display --- (same as before)
    def update_selected_files_display(self):
         try:
             if not self.selected_files_label.winfo_exists() or not self.files_listbox.winfo_exists():
                 return
             count = len(self.selected_files)
             self.selected_files_label.config(text=f"({count}) {SELECTED_FILES_LABEL_TEXT}")
             self.files_listbox.delete(0, tk.END)
             if self.selected_files:
                 for f_path in self.selected_files:
                     self.files_listbox.insert(tk.END, os.path.basename(f_path))
                 self.files_listbox.config(justify=tk.RIGHT)
             else:
                 self.files_listbox.insert(tk.END, " (لم يتم تحديد ملفات) ")
                 self.files_listbox.config(justify=tk.CENTER)
         except tk.TclError:
             pass

    # --- Select Output Directory --- (same as before)
    def select_output_dir(self):
        parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
        if not parent_widget.winfo_exists(): return
        try:
            directory = filedialog.askdirectory(title="اختر مجلد الحفظ", parent=parent_widget)
            if directory:
                self.output_directory = directory
                display_path = os.path.join(self.output_directory, OUTPUT_FOLDER_NAME)
                try:
                    if self.selected_output_dir_label.winfo_exists():
                       self.selected_output_dir_label.config(text=f"{display_path} :{SELECTED_OUTPUT_DIR_LABEL_TEXT}")
                except tk.TclError:
                    pass
        except Exception as e:
             print(f"Error during directory selection: {e}", file=sys.stderr)
             try:
                messagebox.showerror(APP_TITLE, f"حدث خطأ أثناء اختيار المجلد:\n{e}", parent=parent_widget)
             except tk.TclError: pass

    # --- Clear All Selections and Logs (File Tab) --- (same as before)
    def clear_file_tab(self):
        """Resets the file selection, output directory, log, and progress bar."""
        try:
            if self.worker_thread and self.worker_thread.is_alive():
                 if self.file_tab.winfo_exists():
                    messagebox.showwarning(APP_TITLE, "لا يمكن المسح أثناء معالجة الملفات.", parent=self.file_tab)
                 return
            self.selected_files = []
            if self.files_listbox.winfo_exists():
                self.update_selected_files_display() # Updates label and listbox
            self.output_directory = ""
            if self.selected_output_dir_label.winfo_exists():
                self.selected_output_dir_label.config(text=f"- :{SELECTED_OUTPUT_DIR_LABEL_TEXT}")
            if self.status_label.winfo_exists():
                self.status_label.config(text=READY_STATUS)
            if self.progress_bar.winfo_exists():
                self.progress_bar['value'] = 0
            if self.log_area.winfo_exists():
                self.log_area.config(state=tk.NORMAL, cursor='arrow')
                self.log_area.delete("1.0", tk.END)
                self.log_area.config(state=tk.DISABLED, cursor='arrow')
            if self.process_files_button.winfo_exists():
                self.process_files_button.config(state=tk.NORMAL)
            if self.clear_files_button.winfo_exists():
                self.clear_files_button.config(state=tk.NORMAL)
        except tk.TclError:
             try:
                 if self.file_tab.winfo_exists():
                     messagebox.showerror(APP_TITLE, "حدث خطأ أثناء مسح تحديدات الملفات.", parent=self.file_tab)
             except tk.TclError: pass
        except Exception as e:
             try:
                 if self.file_tab.winfo_exists():
                     messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع أثناء المسح:\n{e}", parent=self.file_tab)
             except tk.TclError: pass

    # --- Start File Processing --- (same as before, worker handles logic)
    def start_file_processing(self):
        """Validates inputs and starts the file processing thread."""
        parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
        if not parent_widget.winfo_exists(): return

        if not self.selected_files:
            messagebox.showerror(APP_TITLE, ERROR_SELECT_FILES, parent=parent_widget)
            return
        if not self.output_directory:
             messagebox.showerror(APP_TITLE, ERROR_SELECT_OUTPUT_DIR, parent=parent_widget)
             return
        if self.worker_thread and self.worker_thread.is_alive():
             messagebox.showwarning(APP_TITLE, "المعالجة جارية بالفعل.", parent=parent_widget)
             return

        try:
            if self.process_files_button.winfo_exists():
               self.process_files_button.config(state=tk.DISABLED)
            if self.clear_files_button.winfo_exists():
               self.clear_files_button.config(state=tk.DISABLED)
            if self.status_label.winfo_exists():
               self.status_label.config(text=PROCESSING_STATUS)
            if self.progress_bar.winfo_exists():
               self.progress_bar['value'] = 0

            if self.log_area.winfo_exists():
                self.log_area.config(state=tk.NORMAL, cursor='arrow')
                self.log_area.delete("1.0", tk.END)
                self.log_area.insert(tk.END, f"بدء معالجة {len(self.selected_files)} ملف...\n", ("log_base", "info"))
                self.log_area.see(tk.END)
                self.log_area.config(state=tk.DISABLED, cursor="arrow")

            self.worker_thread = threading.Thread(
                target=file_processing_worker,
                args=(self.selected_files, self.output_directory, self.progress_queue),
                daemon=True
            )
            self.worker_thread.start()
        except tk.TclError:
             try:
                 if parent_widget.winfo_exists():
                     messagebox.showerror(APP_TITLE, "حدث خطأ أثناء تحديث واجهة المستخدم لبدء المعالجة.", parent=parent_widget)
             except tk.TclError: pass
             try:
                 if self.process_files_button.winfo_exists():
                     self.process_files_button.config(state=tk.NORMAL)
                 if self.clear_files_button.winfo_exists():
                     self.clear_files_button.config(state=tk.NORMAL)
                 if self.status_label.winfo_exists():
                     self.status_label.config(text=READY_STATUS)
                 if self.log_area.winfo_exists():
                     self.log_area.config(state=tk.DISABLED, cursor="arrow")
             except tk.TclError:
                 pass

    # --- Log Message --- (same as before)
    def log_message(self, message, tags):
        try:
            if not self.log_area.winfo_exists(): return
            self.log_area.config(state=tk.NORMAL, cursor='arrow')
            effective_tags = ("log_base",) + (tags if isinstance(tags, tuple) else (tags,))
            self.log_area.insert(tk.END, message + "\n", effective_tags)
            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED, cursor='arrow')
        except tk.TclError as e:
            print(f"Error writing to log area (widget likely destroyed): {e}", file=sys.stderr)
        except Exception as e:
            print(f"Unexpected error in log_message: {e}", file=sys.stderr)
            try:
                 if self.log_area.winfo_exists():
                     self.log_area.config(state=tk.DISABLED, cursor='arrow')
            except tk.TclError:
                 pass

    # --- Check Progress Queue --- (same as before)
    def check_progress_queue(self):
        try:
            if not self.root.winfo_exists(): return
            while True:
                message_type, data = self.progress_queue.get_nowait()
                if not self.root.winfo_exists(): return

                if message_type == "progress":
                    if self.progress_bar.winfo_exists():
                        self.progress_bar['value'] = data
                elif message_type == "log":
                    self.log_message(data, "success")
                elif message_type == "error":
                     self.log_message(data, "error")
                elif message_type == "done":
                    parent_widget = self.file_tab if self.file_tab.winfo_exists() else None
                    if self.status_label.winfo_exists():
                        self.status_label.config(text=DONE_STATUS)
                    if self.process_files_button.winfo_exists():
                        self.process_files_button.config(state=tk.NORMAL)
                    if self.clear_files_button.winfo_exists():
                        self.clear_files_button.config(state=tk.NORMAL)
                    if self.progress_bar.winfo_exists():
                         if self.progress_bar['value'] < 100:
                             self.progress_bar['value'] = 100
                    self.log_message("--- اكتملت المعالجة ---", "info")
                    if parent_widget and parent_widget.winfo_exists():
                        messagebox.showinfo(APP_TITLE, DONE_STATUS, parent=parent_widget)
                    self.worker_thread = None
                    break
        except queue.Empty:
            pass
        except tk.TclError as e:
            print(f"Error updating GUI from queue (widget likely destroyed): {e}", file=sys.stderr)
            return
        except Exception as e:
            print(f"Unexpected error updating GUI from queue: {e}", file=sys.stderr)
            try:
                self.log_message(f"خطأ في تحديث الواجهة: {e}", "error")
            except Exception:
                pass
        if self.root.winfo_exists():
            self.root.after(100, self.check_progress_queue)

    # --- Show About --- (Updated version in text constant)
    def show_about(self):
        parent_widget = self.root if self.root.winfo_exists() else None
        if parent_widget:
            # ABOUT_TEXT constant was already updated with new version 1.2
            messagebox.showinfo(ABOUT_TITLE, ABOUT_TEXT, parent=parent_widget)

    # --- Show Help --- (Updated help text constant)
    def show_help(self):
        if not self.root.winfo_exists(): return
        try:
            help_win = tb.Toplevel(master=self.root, title=HELP_TITLE)
            help_win.transient(self.root)
            help_win.grab_set()
            help_win.geometry("550x530") # Increased height slightly more for XML help
            help_win.resizable(False, False)
            help_frame = ttk.Frame(help_win, padding=10)
            help_frame.pack(expand=True, fill=BOTH)
            help_text_widget = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD, padx=5, pady=5, font=self.arabic_ui_font)
            help_text_widget.pack(expand=True, fill=BOTH, pady=(0, 10))

            # HELP_TEXT constant was already updated
            updated_help_text = HELP_TEXT
            help_text_widget.insert(tk.END, updated_help_text)
            help_text_widget.tag_configure("right", justify='right')
            help_text_widget.tag_add("right", "1.0", "end")
            help_text_widget.config(state=tk.DISABLED, cursor="arrow")
            self.add_context_menu(help_text_widget)

            close_button = tb.Button(help_frame, text="إغلاق", command=help_win.destroy, bootstyle="secondary")
            close_button.pack()

            help_win.update_idletasks()
            if self.root.winfo_exists():
                root_x = self.root.winfo_x(); root_y = self.root.winfo_y()
                root_w = self.root.winfo_width(); root_h = self.root.winfo_height()
                win_w = help_win.winfo_width(); win_h = help_win.winfo_height()
                x = root_x + (root_w // 2) - (win_w // 2)
                y = root_y + (root_h // 2) - (win_h // 2)
                help_win.geometry(f"+{x}+{y}")
            else:
                help_win.geometry(f"+100+100")
            help_win.wait_window()
        except tk.TclError:
            try:
                if self.root.winfo_exists():
                   messagebox.showerror(APP_TITLE, "لم يتمكن من فتح نافذة المساعدة.", parent=self.root)
            except tk.TclError: pass
        except Exception as e:
            try:
                if self.root.winfo_exists():
                    messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع عند فتح المساعدة:\n{e}", parent=self.root)
            except tk.TclError: pass

    # --- Context Menu for Text Widgets ---
    # ... (Keep add_context_menu, show_context_menu, copy_from_disabled_text, copy_listbox_selection, select_all methods as they are) ...
    def add_context_menu(self, widget):
        menu = tk.Menu(widget, tearoff=0, font=self.menu_font_spec)
        is_text_or_scrolled = isinstance(widget, (tk.Text, scrolledtext.ScrolledText))
        is_listbox = isinstance(widget, tk.Listbox)
        is_editable = False; can_select_all = False; can_copy = False
        try:
            if is_text_or_scrolled:
                is_editable = widget.cget('state') == tk.NORMAL
                can_select_all = True; can_copy = True
            elif is_listbox:
                can_copy = True
        except tk.TclError: pass

        if is_editable:
            menu.add_command(label=CUT_TEXT, command=lambda w=widget: w.event_generate("<<Cut>>"), accelerator="Ctrl+X", state=tk.DISABLED)
            menu.add_command(label=COPY_TEXT, command=lambda w=widget: w.event_generate("<<Copy>>"), accelerator="Ctrl+C", state=tk.DISABLED)
            menu.add_command(label=PASTE_TEXT, command=lambda w=widget: w.event_generate("<<Paste>>"), accelerator="Ctrl+V", state=tk.NORMAL)
        elif can_copy:
            if is_listbox: menu.add_command(label=COPY_TEXT, command=lambda w=widget: self.copy_listbox_selection(w), accelerator="Ctrl+C", state=tk.DISABLED)
            else: menu.add_command(label=COPY_TEXT, command=lambda w=widget: self.copy_from_disabled_text(w), accelerator="Ctrl+C", state=tk.DISABLED)
        if can_select_all:
             if menu.index(tk.END) is not None: menu.add_separator()
             menu.add_command(label=SELECT_ALL_TEXT, command=lambda w=widget: self.select_all(w), accelerator="Ctrl+A", state=tk.DISABLED)
        if menu.index(tk.END) is not None:
            popup_button = "<Button-2>" if sys.platform == "darwin" else "<Button-3>"
            widget.bind(popup_button, lambda event, m=menu, w=widget: self.show_context_menu(event, m, w))

    def show_context_menu(self, event, menu, widget):
        try:
            if not widget.winfo_exists() or not menu.winfo_exists(): return
            has_selection = False
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                if widget.tag_ranges(tk.SEL): has_selection = True
            elif isinstance(widget, tk.Listbox):
                 if widget.curselection(): has_selection = True
            is_editable = False
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                 try: is_editable = widget.cget('state') == tk.NORMAL
                 except tk.TclError: pass
            for item_label in [CUT_TEXT, COPY_TEXT, PASTE_TEXT, SELECT_ALL_TEXT]:
                try:
                    item_index = menu.index(item_label)
                    if item_index is not None:
                        new_state = tk.DISABLED
                        if item_label == CUT_TEXT: new_state = tk.NORMAL if has_selection and is_editable else tk.DISABLED
                        elif item_label == COPY_TEXT: new_state = tk.NORMAL if has_selection else tk.DISABLED
                        elif item_label == PASTE_TEXT: new_state = tk.NORMAL if is_editable else tk.DISABLED
                        elif item_label == SELECT_ALL_TEXT:
                             if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                                 try:
                                     has_content = widget.get("1.0", "end-1c").strip()
                                     new_state = tk.NORMAL if has_content else tk.DISABLED
                                 except tk.TclError: pass
                             else: new_state = tk.DISABLED
                        menu.entryconfig(item_index, state=new_state)
                except tk.TclError: continue
            menu.tk_popup(event.x_root, event.y_root)
        except tk.TclError: pass

    def copy_from_disabled_text(self, widget):
        if not isinstance(widget, (tk.Text, scrolledtext.ScrolledText)): return
        try:
            if not widget.winfo_exists() or not self.root.winfo_exists(): return
            if widget.tag_ranges(tk.SEL):
                selected_text = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.root.update()
        except tk.TclError as e: print(f"Error copying from disabled text widget: {e}", file=sys.stderr)

    def copy_listbox_selection(self, listbox):
        try:
             if not listbox.winfo_exists() or not self.root.winfo_exists(): return
             selected_indices = listbox.curselection()
             if not selected_indices: return
             selected_items = [listbox.get(i) for i in selected_indices]
             text_to_copy = "\n".join(selected_items)
             self.root.clipboard_clear()
             self.root.clipboard_append(text_to_copy)
             self.root.update()
        except tk.TclError as e: print(f"Error copying from listbox: {e}", file=sys.stderr)

    def select_all(self, widget):
        if not isinstance(widget, (tk.Text, scrolledtext.ScrolledText)): return 'break'
        if not widget.winfo_exists(): return 'break'
        original_state = None
        try:
            if widget.cget('state') == tk.DISABLED:
                original_state = tk.DISABLED
                widget.config(state=tk.NORMAL)
            widget.tag_remove(tk.SEL, "1.0", tk.END)
            widget.tag_add(tk.SEL, "1.0", "end-1c")
            widget.mark_set(tk.INSERT, "1.0")
            widget.see(tk.INSERT)
        except tk.TclError: pass
        finally:
             try:
                 if original_state == tk.DISABLED and widget.winfo_exists(): widget.config(state=tk.DISABLED)
             except tk.TclError: pass
        return 'break'

    # --- Window Closing Handler --- (same as before)
    def on_closing(self):
        if not self.root.winfo_exists(): return
        try:
            if self.worker_thread and self.worker_thread.is_alive():
               if messagebox.askokcancel("خروج", "معالجة الملفات لا تزال قيد التشغيل. هل تريد الخروج على أي حال؟\n(قد لا تكتمل معالجة الملف الحالي)", parent=self.root):
                   self.root.destroy()
               else:
                   return
            else:
               self.root.destroy()
        except tk.TclError:
            try:
                if self.root.winfo_exists(): self.root.destroy()
            except tk.TclError: pass
        except Exception as e:
            print(f"Unexpected error during closing: {e}", file=sys.stderr)
            try:
                if self.root.winfo_exists(): self.root.destroy()
            except tk.TclError: pass

# --- Main Execution --- (same as before)
if __name__ == "__main__":
    try:
        if sys.stdout.encoding is None or sys.stdout.encoding.lower() != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')
        if sys.stderr.encoding is None or sys.stderr.encoding.lower() != 'utf-8': sys.stderr.reconfigure(encoding='utf-8')
    except Exception as e:
         print(f"Note: Could not reconfigure stdout/stderr encoding to UTF-8: {e}", file=sys.stderr)
    try:
        root = tb.Window(themename="litera")
        app = ArabicProcessorApp(root)
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        root.mainloop()
    except tk.TclError as e:
        print(f"Fatal Tkinter error on startup: {e}", file=sys.stderr)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Fatal Tkinter error:\n{e}\n\nThe application cannot start.", "Application Error", 0x10)
        except Exception: pass
    except Exception as e:
         print(f"Fatal error during application startup: {e}", file=sys.stderr)
