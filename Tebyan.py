import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, font as tkfont, simpledialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import arabic_reshaper # Import the library
from bidi.algorithm import get_display
import sys
import os
import json
import csv
import xml.etree.ElementTree as ET # Import ElementTree
from xml.etree.ElementTree import ParseError as ETParseError # Import specific ParseError
import threading
import queue # For thread-safe GUI updates
import re # For splitting key=value and for the new core processing
import textwrap # For the new max-length splitting

# --- Try to import tkinterdnd2 for drag-and-drop ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD, COPY as DND_ACTION_COPY, CF_HDROP
    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False
    # These messages will appear in the console if tkinterdnd2 is not installed
    print("مكتبة tkinterdnd2 غير موجودة. خاصية سحب وإفلات الملفات في تبويب معالجة الملفات ستكون معطلة.", file=sys.stderr)
    print("لتمكين هذه الخاصية، يرجى تثبيت المكتبة باستخدام الأمر التالي في الطرفية أو موجه الأوامر:", file=sys.stderr)
    print("pip install tkinterdnd2", file=sys.stderr)
    print("-" * 60, file=sys.stderr)

##--NEW--##
# --- Try to import pyyaml for YAML support ---
try:
    import yaml
    from yaml.error import YAMLError
    YAML_SUPPORTED = True
except ImportError:
    YAML_SUPPORTED = False
    print("مكتبة PyYAML غير موجودة. خاصية معالجة ملفات YAML ستكون معطلة.", file=sys.stderr)
    print("لتمكين هذه الخاصية، يرجى تثبيت المكتبة باستخدام الأمر التالي:", file=sys.stderr)
    print("pip install PyYAML", file=sys.stderr)
    print("-" * 60, file=sys.stderr)


# --- Constants and Arabic UI Strings ---
APP_TITLE = "برنامج تبيان"
# ... (most constants are unchanged) ...
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
##--MODIFIED--##
ABOUT_TEXT = f"{APP_TITLE}\n\nالإصدار: 1.9.0\n\nبرنامج لمعالجة النصوص العربية لعرضها بشكل صحيح في التطبيقات والألعاب التي لا تدعم اللغة العربية.\n- تم إصلاح مشكلة عكس الأكواد مثل '\\n'.\n- تمت إضافة دعم لملفات YAML.\n- تمت إضافة خيار محاذاة النص (يمين/يسار).\n\n MrGamesKingPro Ⓒ 2025  جميع الحقوق محفوظة \n\n  https://github.com/MrGamesKingPro"

##--MODIFIED--##
HELP_TEXT = """
كيفية الاستخدام:

1.  **معالجة نصوص (التبويب الأول):**
    *   اختر محاذاة النص (يمين أو يسار) من الخيارات الجديدة.
    *   اكتب أو الصق النص في المربع العلوي.
    *   اضغط على زر 'معالجة النص'.
    *   سيظهر النص المعدل في المربع السفلي، جاهزاً للنسخ.

2.  **معالجة ملفات (التبويب الثاني):**
    *   اضغط على 'فتح ملفًا أو أكثر' لتحديد ملفات نصية (.txt, .json, .csv, .xml, .yaml).
    *   اضغط على 'مسار المجلد الحفظ' لتحديد مكان حفظ الملفات المعالجة.
    *   اضغط على 'معالجة الملفات المحددة'.

3.  **الخيارات المتقدمة (يمكن الوصول إليها من كلا التبويبين عبر زر "خيارات متقدمة..."):**
    *   **تجاهل الأكواد المخصصة (مهم جدًا):**
        *   عند تفعيل هذا الخيار، يمكنك كتابة قائمة من الأكواد أو الوسوم (Tags) التي تريد حمايتها من المعالجة.
        *   **لحل مشكلة عكس الأكواد مثل `\\n` أو الوسوم مثل `<li>`، يجب عليك إضافتها هنا.**
        *   مثال: `<br>, <li>, \\n, [PLAYER], </color>`
        *   سيقوم البرنامج بتجاهل هذه الأكواد تمامًا أثناء عملية المعالجة، مع الحفاظ عليها في مكانها الصحيح في النص النهائي.

    *   **المعالجة المشروطة حسب الكلمة:**
        *   لن تتم معالجة السطر إلا إذا كان يحتوي على الكلمة أو العبارة التي تحددها.

    *   **تمكين تقسيم الأسطر المتقدم:** عند التفعيل، سيتم تطبيق أحد وضعي التقسيم على النص العربي فقط:
        *   **وضع "تقسيم حسب عدد الكلمات من النهاية":** يقسم النص إلى جزأين.
        *   **وضع "تقسيم حسب طول السطر الأقصى":** يقسم النص إلى أجزاء متعددة، كل جزء لا يتجاوز الطول المحدد تقريبًا.

    *   **فاصل الأجزاء:** النص الذي سيتم إدراجه بين الأجزاء المقسمة (يمكنك الاختيار من القائمة أو كتابة فاصل مخصص).

ملاحظات عامة:
*   في ملفات TXT التي تحتوي على key=value، ستتم معالجة القيمة (value) فقط.
*   في ملفات CSV و JSON و XML و YAML، ستتم محاولة معالجة كل القيم النصية التي تحتوي على حروف عربية.
*   يمكنك استخدام زر الفأرة الأيمن للقص/النسخ/اللصق.
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
OUTPUT_FOLDER_NAME = "processed_files"

ADVANCED_OPTIONS_BUTTON_TEXT = "خيارات متقدمة..."
ADVANCED_SPLIT_DIALOG_TITLE = "إعدادات المعالجة المتقدمة"
ENABLE_WORD_FILTER_CHECKBOX_TEXT = "تمكين المعالجة المشروطة (فقط للأسطر التي تحتوي على كلمة معينة)"
FILTER_WORD_LABEL_TEXT = ":الكلمة/العبارة المشروطة"
ENABLE_SPLITTING_CHECKBOX_TEXT = "تمكين تقسيم الأسطر المتقدم"
SPLIT_MODE_LABEL_TEXT = ":وضع التقسيم"
WORDS_FROM_END_MODE_TEXT = "تقسيم حسب عدد الكلمات من النهاية"
MAX_LENGTH_MODE_TEXT = "تقسيم حسب طول السطر الأقصى"
WORDS_FROM_END_LABEL_TEXT = ":عدد الكلمات من نهاية السطر للجزء الأول"
MAX_CHARS_LABEL_TEXT = ":الحد الأقصى لعدد الأحرف في السطر الواحد"
SEPARATOR_LABEL_TEXT = ":فاصل الأجزاء"
SEPARATOR_COMBOBOX_LABEL_TEXT = ":اختر فاصلًا أو اكتب مخصصًا"
ERROR_INVALID_WORD_COUNT = "عدد الكلمات يجب أن يكون رقمًا صحيحًا (0 أو أكبر)."
ERROR_INVALID_MAX_LENGTH = "الحد الأقصى لعدد الأحرف يجب أن يكون رقمًا صحيحًا أكبر من 0."
OK_BUTTON_TEXT = "موافق"
CANCEL_BUTTON_TEXT = "إلغاء"

ENABLE_IGNORE_CODES_CHECKBOX_TEXT = "تمكين تجاهل الأكواد المخصصة أثناء المعالجة"
IGNORE_CODES_LABEL_TEXT = ":الأكواد المراد تجاهلها (يفصل بينها بفاصلة ,)"
##--MODIFIED--##
DEFAULT_IGNORE_CODES = "<br>, </br>, <color=...>, </color>, [ICON], \\n"

PREDEFINED_SEPARATORS = { # Display name: actual value
    "سطر جديد (\\n)": "\n",
    "وسم HTML (<br>)": "<br>",
    "وسم BR ثم سطر جديد (<br>\\n)": "<br>\n",
    "فاصل رأسي ( | )": " | ",
    "مسافة فقط ( )": " ",
    "فاصل أفقي (---)": "---",
}

##--NEW--##
# --- New UI Strings for Alignment ---
ALIGNMENT_LABEL_TEXT = ":محاذاة النص"
ALIGN_RIGHT_TEXT = "يمين"
ALIGN_LEFT_TEXT = "يسار"


# --- Helper to check for Arabic characters ---
def _is_arabic(text_segment):
    if not isinstance(text_segment, str) or not text_segment:
        return False
    # This check is efficient for single characters as well
    return any('\u0600' <= char <= '\u06FF' or
               '\u0750' <= char <= '\u077F' or
               '\u08A0' <= char <= '\u08FF' or
               '\uFB50' <= char <= '\uFDFF' or
               '\uFE70' <= char <= '\uFEFF'
               for char in text_segment)

# --- Core Processing Function (unchanged) ---
def process_arabic_text_core(input_text):
    if not _is_arabic(input_text):
        return input_text
    try:
        configuration = {'delete_harakat': False, 'support_ligatures': True}
        reshaper_instance = arabic_reshaper.ArabicReshaper(configuration=configuration)
        reshaped_text = reshaper_instance.reshape(input_text)
        bidi_text = get_display(reshaped_text)
        return bidi_text
    except Exception as e:
        print(f"Error in process_arabic_text_core for: '{input_text[:50]}...' - {e}", file=sys.stderr)
        return input_text

# --- Text Splitting Utility Functions (unchanged) ---
def split_text_into_two_parts(text, words_for_first_part_from_end):
    if not isinstance(text, str) or not text.strip():
        return text, ""
    words = text.split()
    num_words = len(words)
    if words_for_first_part_from_end <= 0:
        return "", text
    if words_for_first_part_from_end >= num_words:
        return text, ""
    split_idx_for_second_part_end = num_words - words_for_first_part_from_end
    first_segment_words = words[split_idx_for_second_part_end:]
    second_segment_words = words[:split_idx_for_second_part_end]
    first_segment = " ".join(first_segment_words)
    second_segment = " ".join(second_segment_words)
    return first_segment, second_segment

def split_string_by_length_with_word_awareness(text, max_chars):
    if not isinstance(text, str) or not text.strip() or max_chars <= 0:
        return [text] # Return as a list with one item
    wrapped_lines = textwrap.wrap(text, width=max_chars, break_long_words=True, break_on_hyphens=False, replace_whitespace=False, drop_whitespace=True)
    return wrapped_lines if wrapped_lines else [text]

# --- File Processing Logic ---
def process_txt_file(input_path, output_path, transform_function):
    try:
        with open(input_path, 'r', encoding='utf-8') as infile, \
             open(output_path, 'w', encoding='utf-8') as outfile:
            for line in infile:
                original_line_rstrip = line.rstrip('\r\n')
                
                parts = original_line_rstrip.split('=', 1)
                if len(parts) == 2:
                    key_part, value_part = parts
                    processed_value = transform_function(value_part)
                    outfile.write(key_part + "=" + processed_value + '\n')
                else:
                    transformed_whole_line = transform_function(original_line_rstrip)
                    outfile.write(transformed_whole_line + '\n')
        return True, None
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}: {e}"

def process_csv_file(input_path, output_path, transform_function):
    try:
        rows = []
        with open(input_path, 'r', encoding='utf-8', newline='') as infile:
            reader = csv.reader(infile)
            try:
                header = next(reader)
                processed_header = [transform_function(cell) if isinstance(cell, str) else cell for cell in header]
                rows.append(processed_header)
            except StopIteration:
                 return True, None 
            for row in reader:
                processed_row = [transform_function(cell) if isinstance(cell, str) else cell for cell in row]
                rows.append(processed_row)
        with open(output_path, 'w', encoding='utf-8', newline='') as outfile:
            writer = csv.writer(outfile)
            writer.writerows(rows)
        return True, None
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (CSV): {e}"

##--MODIFIED--##
# Refactored recursive data processor for use by both JSON and YAML
def _process_data_recursively(value, transform_function):
    if isinstance(value, str):
        return transform_function(value)
    elif isinstance(value, dict):
        return {k: _process_data_recursively(v, transform_function) for k, v in value.items()}
    elif isinstance(value, list):
        return [_process_data_recursively(item, transform_function) for item in value]
    else:
        return value

def process_json_file(input_path, output_path, transform_function):
    try:
        with open(input_path, 'r', encoding='utf-8') as infile:
            data = json.load(infile)
        processed_data = _process_data_recursively(data, transform_function)
        with open(output_path, 'w', encoding='utf-8') as outfile:
            json.dump(processed_data, outfile, ensure_ascii=False, indent=4)
        return True, None
    except json.JSONDecodeError as e:
         return False, f"{ERROR_READING_FILE} (JSON Decode): {e}"
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (JSON): {e}"

##--NEW--##
def process_yaml_file(input_path, output_path, transform_function):
    if not YAML_SUPPORTED:
        return False, "مكتبة PyYAML غير مثبتة."
    try:
        with open(input_path, 'r', encoding='utf-8') as infile:
            data = yaml.safe_load(infile)
        
        processed_data = _process_data_recursively(data, transform_function)
        
        with open(output_path, 'w', encoding='utf-8') as outfile:
            yaml.dump(processed_data, outfile, allow_unicode=True, sort_keys=False)
        return True, None
    except YAMLError as e:
        return False, f"{ERROR_READING_FILE} (YAML Parse): {e}"
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (YAML): {e}"

def process_xml_file(input_path, output_path, transform_function):
    try:
        parser = ET.XMLParser(encoding='utf-8', target=ET.TreeBuilder())
        tree = ET.parse(input_path, parser=parser)
        root = tree.getroot()

        for element in root.iter():
            if element.text and isinstance(element.text, str):
                element.text = transform_function(element.text)
            if element.tail and isinstance(element.tail, str):
                element.tail = transform_function(element.tail)
            for attr_key, attr_value in element.attrib.items():
                if isinstance(attr_value, str):
                    element.attrib[attr_key] = transform_function(attr_value)
        
        tree.write(output_path, encoding='utf-8', xml_declaration=True, method="xml")
        return True, None
    except ETParseError as e:
        return False, f"{ERROR_READING_FILE} (XML Parse): {e}"
    except Exception as e:
        return False, f"{ERROR_READING_FILE}/{ERROR_WRITING_FILE}/{ERROR_PROCESSING_FILE} (XML): {e}"

# --- Worker Thread for File Processing ---
##--MODIFIED--##
def file_processing_worker(file_list, output_dir, progress_queue, text_transformation_func):
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
        if not progress_queue.full():
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
            progress_queue.put(("log", f"معالجة: {base_name} ...")) 
            if ext_lower == '.txt':
                success, error_msg = process_txt_file(input_path, output_path, text_transformation_func)
            elif ext_lower == '.csv':
                 success, error_msg = process_csv_file(input_path, output_path, text_transformation_func)
            elif ext_lower == '.json':
                 success, error_msg = process_json_file(input_path, output_path, text_transformation_func)
            elif ext_lower == '.xml':
                 success, error_msg = process_xml_file(input_path, output_path, text_transformation_func)
            elif ext_lower in ['.yaml', '.yml'] and YAML_SUPPORTED: # Added YAML support
                 success, error_msg = process_yaml_file(input_path, output_path, text_transformation_func)
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
                 progress_queue.put(("progress", progress)) 
        except Exception as e:
            processed_count += 1
            progress = int((processed_count / total_files) * 100) if total_files > 0 else 100
            progress_queue.put(("error", f"Unexpected error processing '{os.path.basename(input_path)}': {e}"))
            progress_queue.put(("progress", progress))

    progress_queue.put(("done", None))

# --- Advanced Split Options Dialog (Unchanged) ---
class AdvancedSplitOptionsDialog(tk.Toplevel):
    def __init__(self, parent, app_instance):
        super().__init__(parent)
        self.parent_app = app_instance
        self.transient(parent)
        self.grab_set()
        self.title(ADVANCED_SPLIT_DIALOG_TITLE)
        self.resizable(False, False)
        try:
            self.style = parent.style 
            self.arabic_ui_font = parent.arabic_ui_font
        except AttributeError:
            self.style = tb.Style(theme="litera") 
            self.arabic_ui_font = tkfont.Font(family="Tahoma", size=10)

        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(expand=True, fill=BOTH)
        
        # --- NEW: Ignore Codes Section ---
        ignore_frame = tb.Labelframe(main_frame, text="تجاهل الأكواد", bootstyle="danger", padding=10)
        ignore_frame.pack(fill=X, pady=(0, 15))

        self.ignore_enabled_var = tk.BooleanVar(value=self.parent_app.ignore_codes_enabled)
        ignore_check = tb.Checkbutton(ignore_frame, text=ENABLE_IGNORE_CODES_CHECKBOX_TEXT, variable=self.ignore_enabled_var, bootstyle="danger-round-toggle", command=self.toggle_ignore_codes_entry_state)
        ignore_check.pack(anchor=E, pady=(0, 10))
        
        ignore_entry_frame = ttk.Frame(ignore_frame)
        ignore_entry_frame.pack(fill=X, expand=True)
        ignore_codes_label = ttk.Label(ignore_entry_frame, text=IGNORE_CODES_LABEL_TEXT, font=self.arabic_ui_font, wraplength=400, justify=RIGHT)
        ignore_codes_label.pack(side=RIGHT, padx=(0, 5), fill=Y)
        self.ignore_codes_var = tk.StringVar(value=self.parent_app.ignore_codes_raw_string)
        self.ignore_codes_entry = tb.Entry(ignore_entry_frame, textvariable=self.ignore_codes_var, width=40, justify=LEFT)
        self.ignore_codes_entry.pack(side=RIGHT, fill=X, expand=True)

        # --- Conditional Word Filter Section ---
        filter_frame = tb.Labelframe(main_frame, text="المعالجة المشروطة", bootstyle="primary", padding=10)
        filter_frame.pack(fill=X, pady=(0, 15))
        
        self.filter_enabled_var = tk.BooleanVar(value=self.parent_app.filter_by_word_enabled)
        filter_check = tb.Checkbutton(filter_frame, text=ENABLE_WORD_FILTER_CHECKBOX_TEXT, variable=self.filter_enabled_var, bootstyle="primary-round-toggle", command=self.toggle_filter_word_entry_state)
        filter_check.pack(anchor=E, pady=(0, 10))
        
        filter_entry_frame = ttk.Frame(filter_frame)
        filter_entry_frame.pack(fill=X, expand=True)
        filter_word_label = ttk.Label(filter_entry_frame, text=FILTER_WORD_LABEL_TEXT, font=self.arabic_ui_font)
        filter_word_label.pack(side=RIGHT, padx=(0, 5))
        self.filter_word_var = tk.StringVar(value=self.parent_app.filter_word)
        self.filter_word_entry = tb.Entry(filter_entry_frame, textvariable=self.filter_word_var, width=30, justify=LEFT)
        self.filter_word_entry.pack(side=RIGHT, fill=X, expand=True)

        # --- Splitting Options Section ---
        split_options_main_frame = tb.Labelframe(main_frame, text="خيارات تقسيم الأسطر", bootstyle="info", padding=10)
        split_options_main_frame.pack(fill=X, expand=True)

        # Enable Checkbox
        self.enable_var = tk.BooleanVar(value=self.parent_app.split_enabled)
        enable_check = tb.Checkbutton(split_options_main_frame, text=ENABLE_SPLITTING_CHECKBOX_TEXT, variable=self.enable_var, bootstyle="info-round-toggle", command=self.toggle_main_options_state)
        enable_check.pack(anchor=E, pady=(0, 15))
        
        self.options_frame = ttk.Frame(split_options_main_frame)
        self.options_frame.pack(fill=X, expand=True)
        
        # Mode Selection
        mode_frame = ttk.Frame(self.options_frame)
        mode_frame.pack(fill=X, pady=(0,10))
        mode_label = ttk.Label(mode_frame, text=SPLIT_MODE_LABEL_TEXT, font=self.arabic_ui_font)
        mode_label.pack(side=RIGHT, padx=(0, 5))
        
        self.split_mode_var = tk.StringVar(value=self.parent_app.split_mode)
        self.words_mode_radio = tb.Radiobutton(mode_frame, text=WORDS_FROM_END_MODE_TEXT, variable=self.split_mode_var, value="words_from_end", command=self.toggle_mode_specific_options_state, bootstyle="primary-toolbutton")
        self.words_mode_radio.pack(side=RIGHT, padx=2, fill=X, expand=True)
        self.length_mode_radio = tb.Radiobutton(mode_frame, text=MAX_LENGTH_MODE_TEXT, variable=self.split_mode_var, value="max_length", command=self.toggle_mode_specific_options_state, bootstyle="primary-toolbutton")
        self.length_mode_radio.pack(side=RIGHT, padx=2, fill=X, expand=True)


        # Words from End Options
        self.words_options_frame = ttk.Frame(self.options_frame)
        self.words_options_frame.pack(fill=X, pady=5)
        words_label = ttk.Label(self.words_options_frame, text=WORDS_FROM_END_LABEL_TEXT, font=self.arabic_ui_font)
        words_label.pack(side=RIGHT, padx=(0, 5))
        self.words_var = tk.StringVar(value=str(self.parent_app.split_words_from_end))
        self.words_entry = tb.Entry(self.words_options_frame, textvariable=self.words_var, width=10, justify=LEFT)
        self.words_entry.pack(side=RIGHT, fill=X, expand=True)

        # Max Length Options
        self.length_options_frame = ttk.Frame(self.options_frame)
        self.length_options_frame.pack(fill=X, pady=5)
        max_chars_label = ttk.Label(self.length_options_frame, text=MAX_CHARS_LABEL_TEXT, font=self.arabic_ui_font)
        max_chars_label.pack(side=RIGHT, padx=(0, 5))
        self.max_len_var = tk.StringVar(value=str(self.parent_app.max_line_length))
        self.max_len_entry = tb.Entry(self.length_options_frame, textvariable=self.max_len_var, width=10, justify=LEFT)
        self.max_len_entry.pack(side=RIGHT, fill=X, expand=True)

        # Separator Options (Common)
        separator_outer_frame = ttk.Frame(self.options_frame)
        separator_outer_frame.pack(fill=X, pady=(10,5))
        separator_label_widget = ttk.Label(separator_outer_frame, text=SEPARATOR_LABEL_TEXT, font=self.arabic_ui_font)
        separator_label_widget.pack(side=RIGHT, padx=(0,5))
        separator_input_frame = ttk.Frame(separator_outer_frame)
        separator_input_frame.pack(side=RIGHT, fill=X, expand=True)
        self.separator_var = tk.StringVar(value=self.parent_app.split_separator_raw)
        self.separator_combobox = tb.Combobox(
            separator_input_frame, 
            textvariable=self.separator_var,
            values=list(PREDEFINED_SEPARATORS.keys()), 
            font=self.arabic_ui_font,
            justify=LEFT,
            state="readonly" # Prevent custom user typing for simplicity now
        )
        current_raw_separator = self.parent_app.split_separator_raw
        display_value_to_set = current_raw_separator
        for display_name, raw_val in PREDEFINED_SEPARATORS.items():
            if raw_val == current_raw_separator:
                display_value_to_set = display_name
                break
        self.separator_var.set(display_value_to_set)
        # If the current value is not in predefined, set to first one
        if self.separator_var.get() not in PREDEFINED_SEPARATORS:
             self.separator_var.set(list(PREDEFINED_SEPARATORS.keys())[0])

        self.separator_combobox.pack(side=RIGHT, fill=X, expand=True)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(15, 0))
        ok_button = tb.Button(button_frame, text=OK_BUTTON_TEXT, command=self.ok_action, bootstyle="success")
        ok_button.pack(side=LEFT, padx=5, expand=True, fill=X)
        cancel_button = tb.Button(button_frame, text=CANCEL_BUTTON_TEXT, command=self.destroy, bootstyle="secondary")
        cancel_button.pack(side=RIGHT, padx=5, expand=True, fill=X)

        self.toggle_ignore_codes_entry_state() # Initial state for new ignore codes
        self.toggle_filter_word_entry_state() # Initial state for new filter
        self.toggle_main_options_state()      # Initial state for splitting
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.center_dialog(parent)

    def center_dialog(self, parent_window):
        self.update_idletasks()
        parent_x = parent_window.winfo_x()
        parent_y = parent_window.winfo_y()
        parent_w = parent_window.winfo_width()
        parent_h = parent_window.winfo_height()
        dialog_w = self.winfo_width()
        dialog_h = self.winfo_height()
        x = parent_x + (parent_w // 2) - (dialog_w // 2)
        y = parent_y + (parent_h // 2) - (dialog_h // 2)
        self.geometry(f"+{x}+{y}")

    def toggle_ignore_codes_entry_state(self):
        is_enabled = self.ignore_enabled_var.get()
        state = tk.NORMAL if is_enabled else tk.DISABLED
        self.ignore_codes_entry.config(state=state)

    def toggle_filter_word_entry_state(self):
        is_enabled = self.filter_enabled_var.get()
        state = tk.NORMAL if is_enabled else tk.DISABLED
        self.filter_word_entry.config(state=state)

    def toggle_main_options_state(self):
        is_enabled = self.enable_var.get()
        main_state = tk.NORMAL if is_enabled else tk.DISABLED
        
        self.words_mode_radio.config(state=main_state)
        self.length_mode_radio.config(state=main_state)
        self.separator_combobox.config(state=main_state)

        if not is_enabled:
            self.words_entry.config(state=tk.DISABLED)
            self.max_len_entry.config(state=tk.DISABLED)
        else:
            self.toggle_mode_specific_options_state()

    def toggle_mode_specific_options_state(self):
        if not self.enable_var.get():
            return

        current_mode = self.split_mode_var.get()
        if current_mode == "words_from_end":
            self.words_entry.config(state=tk.NORMAL)
            self.max_len_entry.config(state=tk.DISABLED)
        elif current_mode == "max_length":
            self.words_entry.config(state=tk.DISABLED)
            self.max_len_entry.config(state=tk.NORMAL)
        else:
            self.words_entry.config(state=tk.DISABLED)
            self.max_len_entry.config(state=tk.DISABLED)

    def ok_action(self):
        # Save Ignore Codes settings
        self.parent_app.ignore_codes_enabled = self.ignore_enabled_var.get()
        raw_codes_string = self.ignore_codes_var.get()
        self.parent_app.ignore_codes_raw_string = raw_codes_string
        # Process into a list, sorting by length descending to handle overlapping codes (e.g., '</color>' before '<color>')
        codes = [code.strip() for code in raw_codes_string.replace('\\n', '\n').replace('\\t', '\t').split(',') if code.strip()]
        self.parent_app.ignore_codes_list = sorted(codes, key=len, reverse=True)

        # Save Filter settings
        self.parent_app.filter_by_word_enabled = self.filter_enabled_var.get()
        self.parent_app.filter_word = self.filter_word_var.get().strip()

        # Save Splitting settings
        is_enabled = self.enable_var.get()
        current_mode = self.split_mode_var.get()
        words_str = self.words_var.get()
        max_len_str = self.max_len_var.get()
        selected_separator_display = self.separator_var.get()
        # Get raw value from display name, if not found, it's a custom value (though we disabled this)
        separator_raw_to_store = PREDEFINED_SEPARATORS.get(selected_separator_display, selected_separator_display)

        if is_enabled:
            if current_mode == "words_from_end":
                try:
                    words_count = int(words_str)
                    if words_count < 0:
                        messagebox.showerror(ERROR_INVALID_WORD_COUNT, "عدد الكلمات لا يمكن أن يكون سالبًا.", parent=self)
                        return
                    self.parent_app.split_words_from_end = words_count
                except ValueError:
                    messagebox.showerror(ERROR_INVALID_WORD_COUNT, "الرجاء إدخال رقم صحيح لعدد الكلمات.", parent=self)
                    return
            elif current_mode == "max_length":
                try:
                    max_len = int(max_len_str)
                    if max_len <= 0:
                        messagebox.showerror(ERROR_INVALID_MAX_LENGTH, "الحد الأقصى لعدد الأحرف يجب أن يكون أكبر من صفر.", parent=self)
                        return
                    self.parent_app.max_line_length = max_len
                except ValueError:
                    messagebox.showerror(ERROR_INVALID_MAX_LENGTH, "الرجاء إدخال رقم صحيح للحد الأقصى لعدد الأحرف.", parent=self)
                    return
        
        self.parent_app.split_enabled = is_enabled
        self.parent_app.split_mode = current_mode
        self.parent_app.split_separator_raw = separator_raw_to_store
        self.destroy()


# --- GUI Application Class ---
class ArabicProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        if hasattr(root, 'style'):
            self.style = root.style
        else:
            self.style = tb.Style(theme="litera")

        self.root.minsize(650, 750) # Increased min height for alignment options

        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.arabic_font_family = "Tahoma"
        self.default_font_size = self.default_font.actual()["size"]
        try: tkfont.Font(family=self.arabic_font_family)
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

        # Advanced processing configuration
        self.ignore_codes_enabled = True # Enabled by default now
        self.ignore_codes_raw_string = DEFAULT_IGNORE_CODES
        initial_codes = [code.strip() for code in self.ignore_codes_raw_string.replace('\\n', '\n').replace('\\t', '\t').split(',') if code.strip()]
        self.ignore_codes_list = sorted(initial_codes, key=len, reverse=True)

        self.filter_by_word_enabled = False
        self.filter_word = ""
        self.split_enabled = False
        self.split_mode = "words_from_end" # "words_from_end" or "max_length"
        self.split_words_from_end = 1
        self.max_line_length = 15
        self.split_separator_raw = "\n" 
        
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

        self.notebook = ttk.Notebook(root, bootstyle="primary")
        self.notebook.pack(pady=10, padx=10, expand=True, fill=BOTH)

        self.direct_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.direct_tab, text=DIRECT_TAB_TEXT)
        
        ##--NEW--##
        # --- Text Alignment Controls ---
        self.text_alignment = tk.StringVar(value='right')
        alignment_frame = tb.Labelframe(self.direct_tab, text=ALIGNMENT_LABEL_TEXT, bootstyle="info", padding=(10,5))
        alignment_frame.pack(fill=X, pady=(0, 10))
        
        align_right_radio = tb.Radiobutton(alignment_frame, text=ALIGN_RIGHT_TEXT, variable=self.text_alignment, value='right', command=self.update_text_alignment, bootstyle="info-toolbutton")
        align_right_radio.pack(side=RIGHT, padx=2, fill=X, expand=True)
        align_left_radio = tb.Radiobutton(alignment_frame, text=ALIGN_LEFT_TEXT, variable=self.text_alignment, value='left', command=self.update_text_alignment, bootstyle="info-toolbutton")
        align_left_radio.pack(side=RIGHT, padx=2, fill=X, expand=True)
        
        input_label = ttk.Label(self.direct_tab, text=INPUT_LABEL_TEXT)
        input_label.pack(pady=(0, 5), anchor=E)
        self.input_text_area = scrolledtext.ScrolledText(self.direct_tab, height=8, width=60, wrap=tk.WORD, font=self.text_widget_font_spec)
        self.input_text_area.pack(pady=5, padx=5, expand=True, fill=BOTH)
        self.input_text_area.configure(bd=1, relief="sunken", insertbackground='black')
        
        # Configure alignment tags
        self.input_text_area.tag_configure("right", justify='right')
        self.input_text_area.tag_configure("left", justify='left')
        
        self.input_text_area.bind("<KeyRelease>", self._apply_input_alignment)
        self.input_text_area.bind("<Return>", self.handle_input_enter)
        self.input_text_area.focus()
        self.add_context_menu(self.input_text_area)

        direct_action_frame = ttk.Frame(self.direct_tab)
        direct_action_frame.pack(pady=10, fill=X)
        process_button = tb.Button(direct_action_frame, text=PROCESS_BUTTON_TEXT, command=self.process_direct_text, bootstyle="success")
        process_button.pack(side=RIGHT, padx=(5,0), expand=True, fill=X)
        self.direct_adv_options_button = tb.Button(direct_action_frame, text=ADVANCED_OPTIONS_BUTTON_TEXT, command=self.open_advanced_split_options_dialog, bootstyle="secondary-outline")
        self.direct_adv_options_button.pack(side=LEFT, padx=(0,5), expand=True, fill=X)

        output_label = ttk.Label(self.direct_tab, text=OUTPUT_LABEL_TEXT)
        output_label.pack(pady=(5, 5), anchor=E)
        self.output_text_area = scrolledtext.ScrolledText(self.direct_tab, height=8, width=60, wrap=tk.WORD, font=self.text_widget_font_spec)
        self.output_text_area.configure(bd=1, relief="sunken")

        # Configure alignment tags for output
        self.output_text_area.tag_configure("right_readonly", justify='right')
        self.output_text_area.tag_configure("left_readonly", justify='left')
        
        self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
        self.add_context_menu(self.output_text_area)
        self.output_text_area.pack(pady=5, padx=5, expand=True, fill=BOTH)

        direct_bottom_button_frame = ttk.Frame(self.direct_tab)
        direct_bottom_button_frame.pack(pady=(5, 0), fill=X)
        self.clear_direct_button = tb.Button(direct_bottom_button_frame, text=CLEAR_ALL_TEXT_BUTTON_TEXT, command=self.clear_direct_tab, bootstyle="warning")
        self.clear_direct_button.pack(side=LEFT, padx=(0, 5))
        self.copy_all_button = tb.Button(direct_bottom_button_frame, text=COPY_ALL_BUTTON_TEXT, command=self.copy_all_output, bootstyle="info", state=tk.DISABLED)
        self.copy_all_button.pack(side=RIGHT, padx=(5, 0))

        # Initialize text alignment
        self.update_text_alignment()

        self.file_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.file_tab, text=FILE_TAB_TEXT)
        # ... (rest of the file tab is mostly unchanged)
        # ... (code for file_tab, progress bar, log_area, etc. as before)
        file_select_frame = ttk.Frame(self.file_tab)
        file_select_frame.pack(fill=X, pady=5)
        select_files_button = tb.Button(file_select_frame, text=SELECT_FILES_BUTTON_TEXT, command=self.select_files, bootstyle="info")
        select_files_button.pack(side=RIGHT, padx=(10, 0))
        self.selected_files_label = ttk.Label(file_select_frame, text=f"({len(self.selected_files)}) {SELECTED_FILES_LABEL_TEXT}")
        self.selected_files_label.pack(side=RIGHT, fill=X, expand=True, padx=(0, 5))

        self.files_listbox_frame = ttk.Frame(self.file_tab)
        self.files_listbox_frame.pack(pady=5, padx=5, fill=X, side=TOP)
        self.files_listbox_scrollbar = ttk.Scrollbar(self.files_listbox_frame, orient=VERTICAL)
        self.files_listbox_scrollbar.pack(side=RIGHT, fill=Y)
        self.files_listbox = tk.Listbox(self.files_listbox_frame, height=5, width=70, yscrollcommand=self.files_listbox_scrollbar.set, font=self.listbox_font_spec, justify=tk.RIGHT, exportselection=False)
        self.files_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.files_listbox_scrollbar.config(command=self.files_listbox.yview)
        self.add_context_menu(self.files_listbox)
        self.update_selected_files_display()

        if DND_SUPPORTED:
            try:
                self.files_listbox.drop_target_register(DND_FILES)
                self.files_listbox.dnd_bind('<<Drop>>', self.handle_file_drop)
                self.files_listbox.dnd_bind('<<DragEnter>>', self.handle_drag_enter_listbox)
                self.files_listbox.dnd_bind('<<DragOver>>', self.handle_drag_over_listbox)
                self.files_listbox.dnd_bind('<<DragLeave>>', self.handle_drag_leave_listbox)
                self.original_listbox_bg = self.files_listbox.cget("background")
            except Exception as e:
                print(f"فشل في تهيئة خاصية السحب والإفلات لـ files_listbox: {e}", file=sys.stderr)

        output_dir_frame = ttk.Frame(self.file_tab)
        output_dir_frame.pack(fill=X, pady=5)
        select_output_dir_button = tb.Button(output_dir_frame, text=SELECT_OUTPUT_DIR_BUTTON_TEXT, command=self.select_output_dir, bootstyle="info")
        select_output_dir_button.pack(side=RIGHT, padx=(10, 0))
        self.selected_output_dir_label = ttk.Label(output_dir_frame, text=f"- :{SELECTED_OUTPUT_DIR_LABEL_TEXT}", wraplength=450, anchor=E, justify=RIGHT)
        self.selected_output_dir_label.pack(side=RIGHT, fill=X, expand=True, padx=(0, 5))

        file_action_button_frame = ttk.Frame(self.file_tab)
        file_action_button_frame.pack(pady=(10,0), fill=X)
        self.clear_files_button = tb.Button(file_action_button_frame, text=CLEAR_ALL_FILES_BUTTON_TEXT, command=self.clear_file_tab, bootstyle="warning")
        self.clear_files_button.pack(side=LEFT, padx=(0, 5), expand=True, fill=X)
        self.process_files_button = tb.Button(file_action_button_frame, text=PROCESS_FILES_BUTTON_TEXT, command=self.start_file_processing, bootstyle="success")
        self.process_files_button.pack(side=RIGHT, padx=(5, 0), expand=True, fill=X)
        
        file_adv_options_frame = ttk.Frame(self.file_tab)
        file_adv_options_frame.pack(pady=(5,10), fill=X)
        self.file_adv_options_button = tb.Button(file_adv_options_frame, text=ADVANCED_OPTIONS_BUTTON_TEXT, command=self.open_advanced_split_options_dialog, bootstyle="secondary-outline")
        self.file_adv_options_button.pack(expand=True, fill=X)

        status_frame = ttk.Frame(self.file_tab)
        status_frame.pack(fill=X, pady=(0, 0))
        self.status_label = ttk.Label(status_frame, text=READY_STATUS, anchor=E)
        self.status_label.pack(side=RIGHT, padx=(5, 0))
        status_label_title = ttk.Label(status_frame, text=STATUS_LABEL_TEXT, anchor=E)
        status_label_title.pack(side=RIGHT)

        self.progress_bar = tb.Progressbar(self.file_tab, mode='determinate', bootstyle="success-striped")
        self.progress_bar.pack(pady=5, padx=5, fill=X)

        log_font_spec = (self.arabic_font_family, self.arabic_ui_font.actual()["size"] - 1)
        self.log_area = scrolledtext.ScrolledText(self.file_tab, height=6, width=70, wrap=tk.WORD, state=tk.DISABLED, font=log_font_spec)
        self.log_area.pack(pady=5, padx=5, fill=BOTH, expand=True)
        self.log_area.tag_configure("error", foreground="red")
        self.log_area.tag_configure("success", foreground="#008000")
        self.log_area.tag_configure("info", foreground="blue")
        self.log_area.tag_configure("log_base", justify='right')
        self.log_area.config(cursor="arrow")
        self.add_context_menu(self.log_area)

        self.check_progress_queue()

    # --- Drag and Drop Event Handlers (Unchanged) ---
    def handle_drag_enter_listbox(self, event):
        if not DND_SUPPORTED: return
        is_file_drag = False
        if hasattr(event, 'types'):
            for t_spec in event.types:
                if t_spec == DND_FILES or (sys.platform == 'win32' and t_spec == CF_HDROP):
                    is_file_drag = True
                    break
        if is_file_drag:
            try:
                if self.files_listbox.winfo_exists():
                    self.files_listbox.configure(background='lightblue') 
            except tk.TclError: pass 
            return DND_ACTION_COPY 
        return None

    def handle_drag_over_listbox(self, event):
        if not DND_SUPPORTED: return
        is_file_drag = False
        if hasattr(event, 'types'):
            for t_spec in event.types:
                if t_spec == DND_FILES or (sys.platform == 'win32' and t_spec == CF_HDROP):
                    is_file_drag = True
                    break
        if is_file_drag:
            return DND_ACTION_COPY
        return None

    def handle_drag_leave_listbox(self, event):
        if not DND_SUPPORTED: return
        try:
            if hasattr(self, 'original_listbox_bg') and self.files_listbox.winfo_exists():
                self.files_listbox.configure(background=self.original_listbox_bg) 
        except tk.TclError: pass

    def handle_file_drop(self, event):
        if not DND_SUPPORTED: return
        try:
            if hasattr(self, 'original_listbox_bg') and self.files_listbox.winfo_exists():
                self.files_listbox.configure(background=self.original_listbox_bg)
            if not event.data: return
            try:
                files_to_add_raw = event.widget.tk.splitlist(event.data)
            except tk.TclError as e:
                print(f"Error splitting dropped data: {e}. Data: '{event.data}'", file=sys.stderr)
                self.log_message(f"خطأ في تحليل البيانات المسقطة: {event.data}", "error")
                return
            added_count = 0
            current_file_paths_lower = {f.lower() for f in self.selected_files}
            for f_path_raw in files_to_add_raw:
                f_path = os.path.normpath(f_path_raw)
                if os.path.isfile(f_path):
                    if f_path.lower() not in current_file_paths_lower:
                        self.selected_files.append(f_path)
                        current_file_paths_lower.add(f_path.lower()) 
                        added_count += 1
                    else:
                        self.log_message(f"الملف مضاف بالفعل (تم تجاهله): {os.path.basename(f_path)}", "info")
                else:
                    self.log_message(f"العنصر المسقط ليس ملفًا صالحًا أو غير موجود: {os.path.basename(f_path_raw)}", "error")
            if added_count > 0:
                self.update_selected_files_display()
                self.log_message(f"تمت إضافة {added_count} ملفات عن طريق السحب والإفلات.", "info")
        except Exception as e:
            self.log_message(f"خطأ أثناء معالجة الملفات المسقطة: {e}", "error")
            print(f"Error handling file drop: {e}\nDropped data was: '{event.data}'", file=sys.stderr)

    # --- Text Processing Logic ---
    def extract_prefix_core_suffix(self, text):
        """Finds the first and last Arabic char, returning (prefix, core, suffix)."""
        if not isinstance(text, str) or not text:
            return "", text, ""

        first_arabic_idx = -1
        last_arabic_idx = -1
        for i, char in enumerate(text):
            if _is_arabic(char):
                if first_arabic_idx == -1:
                    first_arabic_idx = i
                last_arabic_idx = i

        if first_arabic_idx == -1:
            # No arabic characters, so the whole string is prefix/suffix, core is empty.
            return text, "", ""

        prefix = text[:first_arabic_idx]
        core = text[first_arabic_idx : last_arabic_idx + 1]
        suffix = text[last_arabic_idx + 1 :]
        return prefix, core, suffix

    def _process_segment_with_punctuation_handling(self, seg):
        """Processes a single text segment, preserving leading/trailing non-Arabic chars."""
        if not isinstance(seg, str) or not _is_arabic(seg):
            return seg

        i = 0
        while i < len(seg) and not _is_arabic(seg[i]):
            i += 1
        leading_chars = seg[:i]

        j = len(seg) - 1
        while j >= i and not _is_arabic(seg[j]):
            j -= 1
        trailing_chars = seg[j+1:]
        
        core_seg = seg[i:j+1]

        if not core_seg:
            return seg
            
        processed_core = process_arabic_text_core(core_seg)
        return leading_chars + processed_core + trailing_chars

    ##--MODIFIED--##
    def apply_text_transformations(self, text_input_line):
        """The main text processing pipeline with corrected code handling."""
        if not isinstance(text_input_line, str):
            return text_input_line

        # --- Stage 0: Conditional Word Filter ---
        if self.filter_by_word_enabled and self.filter_word:
            if self.filter_word not in text_input_line:
                return text_input_line

        # --- Stage 1: Placeholder Substitution for Ignored Codes ---
        restoration_map = {}
        line_to_process = text_input_line
        if self.ignore_codes_enabled and self.ignore_codes_list:
            temp_line = line_to_process
            for i, code in enumerate(self.ignore_codes_list):
                # Only perform replacement if the code is actually in the line
                if code in temp_line:
                    placeholder = f"__IGN{i}__"
                    restoration_map[placeholder] = code
                    temp_line = temp_line.replace(code, placeholder)
            line_to_process = temp_line

        # --- Stage 2: Isolate and Process Core Arabic Text ---
        prefix, core_text, suffix = self.extract_prefix_core_suffix(line_to_process)
        
        processed_core = ""
        if not core_text:
            processed_line = line_to_process
        else:
            if not self.split_enabled:
                processed_core = self._process_segment_with_punctuation_handling(core_text)
            else: # Advanced splitting logic
                if self.split_mode == "words_from_end":
                    words_for_first = int(self.split_words_from_end)
                    seg1, seg2 = split_text_into_two_parts(core_text, words_for_first)
                    proc1 = self._process_segment_with_punctuation_handling(seg1)
                    proc2 = self._process_segment_with_punctuation_handling(seg2)
                    processed_core = f"{proc1}{self.split_separator_raw}{proc2}" if proc1 and proc2 else (proc1 or proc2)
                elif self.split_mode == "max_length":
                    max_len = int(self.max_line_length)
                    segments = split_string_by_length_with_word_awareness(core_text, max_len)
                    processed_segments = [self._process_segment_with_punctuation_handling(s) for s in segments]
                    processed_core = self.split_separator_raw.join(processed_segments)
                else: # Fallback
                    processed_core = self._process_segment_with_punctuation_handling(core_text)

            processed_line = prefix + processed_core + suffix

        # --- Stage 2.5: Un-reverse any reversed placeholders ---
        # This makes sure all placeholders are back in their original, non-reversed form
        unreversed_line = processed_line
        if self.ignore_codes_enabled and restoration_map:
            for placeholder in restoration_map.keys():
                reversed_placeholder = placeholder[::-1]
                if reversed_placeholder in unreversed_line:
                    unreversed_line = unreversed_line.replace(reversed_placeholder, placeholder)
        
        # --- Stage 3: Restore Original Codes from Placeholders ---
        final_line = unreversed_line
        if self.ignore_codes_enabled and restoration_map:
            # Sort placeholders by the length of the code they represent, descending.
            # This is crucial for correctly handling nested codes like '</color>' before '<br>'.
            sorted_placeholders = sorted(restoration_map.keys(), key=lambda p: len(restoration_map[p]), reverse=True)
            for placeholder in sorted_placeholders:
                final_line = final_line.replace(placeholder, restoration_map[placeholder])
                
        return final_line

    def open_advanced_split_options_dialog(self):
        dialog = AdvancedSplitOptionsDialog(self.root, self)
        dialog.wait_window()

    ##--NEW--##
    def update_text_alignment(self):
        """Applies the current alignment setting to text widgets."""
        try:
            if not self.root.winfo_exists(): return
            self._apply_input_alignment()
            self._apply_output_alignment()
        except tk.TclError:
            pass # Widget might be destroyed

    def _apply_input_alignment(self, event=None):
        try:
            if not self.input_text_area.winfo_exists(): return
            current_alignment = self.text_alignment.get()
            
            self.input_text_area.tag_remove("right", "1.0", "end")
            self.input_text_area.tag_remove("left", "1.0", "end")
            
            if current_alignment == 'right':
                self.input_text_area.tag_add("right", "1.0", "end")
            else:
                self.input_text_area.tag_add("left", "1.0", "end")
        except tk.TclError: pass
        
    def _apply_output_alignment(self):
        try:
            if not self.output_text_area.winfo_exists(): return
            current_alignment = self.text_alignment.get()
            
            self.output_text_area.config(state=tk.NORMAL)
            self.output_text_area.tag_remove("right_readonly", "1.0", "end")
            self.output_text_area.tag_remove("left_readonly", "1.0", "end")
            
            if current_alignment == 'right':
                self.output_text_area.tag_add("right_readonly", "1.0", "end")
            else:
                self.output_text_area.tag_add("left_readonly", "1.0", "end")
            self.output_text_area.config(state=tk.DISABLED)
        except tk.TclError: pass

    def handle_input_enter(self, event=None):
        try:
            if not self.input_text_area.winfo_exists(): return 'break'
            self.input_text_area.insert(tk.INSERT, '\n')
            self.input_text_area.see(tk.INSERT) 
            self._apply_input_alignment()
            self.input_text_area.focus_set()
            return 'break'
        except tk.TclError: return 'break'
        except Exception as e:
             print(f"Error in handle_input_enter: {e}", file=sys.stderr)
             return 'break'

    def process_direct_text(self):
        has_processed_text = False
        try:
            if not self.input_text_area.winfo_exists() or not self.output_text_area.winfo_exists():
                return
            input_text_block = self.input_text_area.get("1.0", tk.END).strip()
            self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
            self.output_text_area.delete("1.0", tk.END)
            if self.copy_all_button.winfo_exists():
                 self.copy_all_button.config(state=tk.DISABLED)
            if not input_text_block:
                self._apply_input_alignment()
                if self.direct_tab.winfo_exists():
                   messagebox.showwarning(APP_TITLE, "الرجاء إدخال نص للمعالجة.", parent=self.direct_tab)
                self.output_text_area.config(state=tk.DISABLED, cursor='arrow')
                return
            
            lines = input_text_block.splitlines()
            processed_lines = [self.apply_text_transformations(line) for line in lines]
            final_processed_text = "\n".join(processed_lines)
            
            self.output_text_area.insert(tk.END, final_processed_text)
            self._apply_output_alignment() # Apply correct alignment tag
            self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
            has_processed_text = bool(final_processed_text)
            if has_processed_text and self.copy_all_button.winfo_exists():
                self.copy_all_button.config(state=tk.NORMAL)
        except tk.TclError:
             # ... (error handling as before)
             pass
        except Exception as e:
            # ... (error handling as before)
            pass
        finally:
            # ... (finally block as before)
            pass
    
    def clear_direct_tab(self):
        try:
            if self.input_text_area.winfo_exists():
                self.input_text_area.delete("1.0", tk.END)
                self._apply_input_alignment()
            if self.output_text_area.winfo_exists():
                self.output_text_area.config(state=tk.NORMAL, cursor='arrow')
                self.output_text_area.delete("1.0", tk.END)
                self.output_text_area.config(state=tk.DISABLED, cursor="arrow")
            if self.copy_all_button.winfo_exists():
                self.copy_all_button.config(state=tk.DISABLED)
        except tk.TclError:
            # ... (error handling as before)
            pass
        except Exception as e:
            # ... (error handling as before)
            pass
    
    # ... (copy_all_output, save_processed_text are unchanged) ...
    def copy_all_output(self, *args):
        # ... (implementation is unchanged)
        pass
        
    def save_processed_text(self):
        # ... (implementation is unchanged)
        pass

    ##--MODIFIED--##
    def select_files(self):
        parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
        if not parent_widget.winfo_exists(): return
        
        filetypes = [
            ("Text Files", "*.txt"), 
            ("CSV Files", "*.csv"), 
            ("JSON Files", "*.json"), 
            ("XML Files", "*.xml")
        ]
        if YAML_SUPPORTED:
            filetypes.append(("YAML Files", "*.yaml *.yml"))
        filetypes.append(("All Files", "*.*"))
        
        try:
            selected = filedialog.askopenfilenames(
                title="اختر ملفات للمعالجة", 
                filetypes=filetypes,
                parent=parent_widget
            )
            if selected: 
                newly_selected_files = list(selected)
                current_file_paths_lower = {f.lower() for f in self.selected_files}
                added_count = 0
                for f_path in newly_selected_files:
                    if f_path.lower() not in current_file_paths_lower:
                        self.selected_files.append(f_path)
                        current_file_paths_lower.add(f_path.lower())
                        added_count +=1
                if added_count > 0 :
                     self.update_selected_files_display()
                     self.log_message(f"تمت إضافة {added_count} ملفات عبر مربع الحوار.", "info")
        except Exception as e:
             print(f"Error during file selection: {e}", file=sys.stderr)
             try: messagebox.showerror(APP_TITLE, f"حدث خطأ أثناء اختيار الملفات:\n{e}", parent=parent_widget)
             except tk.TclError: pass

    # ... (All other methods remain the same as the original provided code) ...
    # ... (update_selected_files_display, select_output_dir, clear_file_tab, etc.) ...
    # ... (start_file_processing, log_message, check_progress_queue, etc.) ...
    # ... (show_about, show_help, context menus, on_closing) ...
    def update_selected_files_display(self):
         try:
             if not self.selected_files_label.winfo_exists() or not self.files_listbox.winfo_exists(): return
             count = len(self.selected_files)
             self.selected_files_label.config(text=f"({count}) {SELECTED_FILES_LABEL_TEXT}")
             self.files_listbox.delete(0, tk.END)
             if self.selected_files:
                 for f_path in self.selected_files: self.files_listbox.insert(tk.END, os.path.basename(f_path))
                 self.files_listbox.config(justify=tk.RIGHT)
             else:
                 self.files_listbox.insert(tk.END, " (لم يتم تحديد ملفات - يمكنك السحب والإفلات هنا) " if DND_SUPPORTED else " (لم يتم تحديد ملفات) ")
                 self.files_listbox.config(justify=tk.CENTER)
         except tk.TclError: pass

    def select_output_dir(self):
        parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
        if not parent_widget.winfo_exists(): return
        try:
            directory = filedialog.askdirectory(title="اختر مجلد الحفظ", parent=parent_widget)
            if directory:
                self.output_directory = directory
                display_path = os.path.join(self.output_directory, OUTPUT_FOLDER_NAME)
                try:
                    if self.selected_output_dir_label.winfo_exists(): self.selected_output_dir_label.config(text=f"{display_path} :{SELECTED_OUTPUT_DIR_LABEL_TEXT}")
                except tk.TclError: pass
        except Exception as e:
             print(f"Error during directory selection: {e}", file=sys.stderr)
             try: messagebox.showerror(APP_TITLE, f"حدث خطأ أثناء اختيار المجلد:\n{e}", parent=parent_widget)
             except tk.TclError: pass

    def clear_file_tab(self):
        try:
            parent_for_msg = self.file_tab if self.file_tab.winfo_exists() else self.root
            if self.worker_thread and self.worker_thread.is_alive():
                 if parent_for_msg.winfo_exists(): messagebox.showwarning(APP_TITLE, "لا يمكن المسح أثناء معالجة الملفات.", parent=parent_for_msg)
                 return
            self.selected_files = []
            if self.files_listbox.winfo_exists(): self.update_selected_files_display()
            self.output_directory = ""
            if self.selected_output_dir_label.winfo_exists(): self.selected_output_dir_label.config(text=f"- :{SELECTED_OUTPUT_DIR_LABEL_TEXT}")
            if self.status_label.winfo_exists(): self.status_label.config(text=READY_STATUS)
            if self.progress_bar.winfo_exists(): self.progress_bar['value'] = 0
            if self.log_area.winfo_exists():
                self.log_area.config(state=tk.NORMAL, cursor='arrow')
                self.log_area.delete("1.0", tk.END)
                self.log_area.config(state=tk.DISABLED, cursor='arrow')
            if self.process_files_button.winfo_exists(): self.process_files_button.config(state=tk.NORMAL)
            if self.clear_files_button.winfo_exists(): self.clear_files_button.config(state=tk.NORMAL)
        except tk.TclError:
             parent_for_msg = self.file_tab if self.file_tab.winfo_exists() else self.root
             if parent_for_msg.winfo_exists(): messagebox.showerror(APP_TITLE, "حدث خطأ أثناء مسح تحديدات الملفات.", parent=parent_for_msg)
        except Exception as e:
             parent_for_msg = self.file_tab if self.file_tab.winfo_exists() else self.root
             if parent_for_msg.winfo_exists(): messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع أثناء المسح:\n{e}", parent=parent_for_msg)

    def start_file_processing(self):
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
            if self.process_files_button.winfo_exists(): self.process_files_button.config(state=tk.DISABLED)
            if self.clear_files_button.winfo_exists(): self.clear_files_button.config(state=tk.DISABLED)
            if self.status_label.winfo_exists(): self.status_label.config(text=PROCESSING_STATUS)
            if self.progress_bar.winfo_exists(): self.progress_bar['value'] = 0
            if self.log_area.winfo_exists():
                self.log_area.config(state=tk.NORMAL, cursor='arrow')
                self.log_area.delete("1.0", tk.END)
                self.log_area.insert(tk.END, f"بدء معالجة {len(self.selected_files)} ملف...\n", ("log_base", "info"))
                self.log_area.see(tk.END)
                self.log_area.config(state=tk.DISABLED, cursor="arrow")
            self.worker_thread = threading.Thread(target=file_processing_worker, args=(self.selected_files, self.output_directory, self.progress_queue, self.apply_text_transformations), daemon=True)
            self.worker_thread.start()
        except tk.TclError:
             try:
                 if parent_widget.winfo_exists(): messagebox.showerror(APP_TITLE, "حدث خطأ أثناء تحديث واجهة المستخدم لبدء المعالجة.", parent=parent_widget)
             except tk.TclError: pass
             try:
                 if self.process_files_button.winfo_exists(): self.process_files_button.config(state=tk.NORMAL)
                 if self.clear_files_button.winfo_exists(): self.clear_files_button.config(state=tk.NORMAL)
                 if self.status_label.winfo_exists(): self.status_label.config(text=READY_STATUS)
                 if self.log_area.winfo_exists(): self.log_area.config(state=tk.DISABLED, cursor="arrow")
             except tk.TclError: pass

    def log_message(self, message, tags):
        try:
            if not self.log_area.winfo_exists(): return
            self.log_area.config(state=tk.NORMAL, cursor='arrow')
            effective_tags = ("log_base",) + (tags if isinstance(tags, tuple) else (tags,))
            self.log_area.insert(tk.END, message + "\n", effective_tags)
            self.log_area.see(tk.END)
            self.log_area.config(state=tk.DISABLED, cursor='arrow')
        except tk.TclError as e: print(f"Error writing to log area (widget likely destroyed): {e}", file=sys.stderr)
        except Exception as e:
            print(f"Unexpected error in log_message: {e}", file=sys.stderr)
            try:
                 if self.log_area.winfo_exists(): self.log_area.config(state=tk.DISABLED, cursor='arrow')
            except tk.TclError: pass

    def check_progress_queue(self):
        try:
            if not self.root.winfo_exists(): return
            while True:
                message_type, data = self.progress_queue.get_nowait()
                if not self.root.winfo_exists(): return 
                if message_type == "progress":
                    if self.progress_bar.winfo_exists(): self.progress_bar['value'] = data
                elif message_type == "log": self.log_message(data, "success")
                elif message_type == "error": self.log_message(data, "error")
                elif message_type == "done":
                    parent_widget = self.file_tab if self.file_tab.winfo_exists() else self.root
                    if not parent_widget.winfo_exists(): parent_widget = None

                    if self.status_label.winfo_exists(): self.status_label.config(text=DONE_STATUS)
                    if self.process_files_button.winfo_exists(): self.process_files_button.config(state=tk.NORMAL)
                    if self.clear_files_button.winfo_exists(): self.clear_files_button.config(state=tk.NORMAL)
                    if self.progress_bar.winfo_exists():
                         if self.progress_bar['value'] < 100 and len(self.selected_files) > 0 : self.progress_bar['value'] = 100
                    self.log_message("--- اكتملت المعالجة ---", "info")
                    if parent_widget: messagebox.showinfo(APP_TITLE, DONE_STATUS, parent=parent_widget)
                    self.worker_thread = None
                    break 
        except queue.Empty: pass
        except tk.TclError as e:
            print(f"Error updating GUI from queue (widget likely destroyed): {e}", file=sys.stderr)
            return 
        except Exception as e:
            print(f"Unexpected error updating GUI from queue: {e}", file=sys.stderr)
            try: self.log_message(f"خطأ في تحديث الواجهة: {e}", "error")
            except Exception: pass
        if self.root.winfo_exists(): self.root.after(100, self.check_progress_queue)

    def show_about(self):
        parent_widget = self.root if self.root.winfo_exists() else None
        if parent_widget: messagebox.showinfo(ABOUT_TITLE, ABOUT_TEXT, parent=parent_widget)

    def show_help(self):
        if not self.root.winfo_exists(): return
        try:
            help_win = tb.Toplevel(master=self.root, title=HELP_TITLE)
            help_win.transient(self.root); help_win.grab_set()
            help_win.geometry("620x700"); help_win.resizable(False, False) # Adjusted size
            help_frame = ttk.Frame(help_win, padding=10)
            help_frame.pack(expand=True, fill=BOTH)
            help_text_widget = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD, padx=5, pady=5, font=self.arabic_ui_font)
            help_text_widget.pack(expand=True, fill=BOTH, pady=(0,10))
            help_text_widget.insert(tk.END, HELP_TEXT)
            help_text_widget.tag_configure("right", justify='right')
            help_text_widget.tag_add("right", "1.0", "end")
            help_text_widget.config(state=tk.DISABLED, cursor="arrow")
            self.add_context_menu(help_text_widget)
            close_button = tb.Button(help_frame, text="إغلاق", command=help_win.destroy, bootstyle="secondary")
            close_button.pack()
            help_win.update_idletasks()
            if self.root.winfo_exists():
                root_x, root_y = self.root.winfo_x(), self.root.winfo_y()
                root_w, root_h = self.root.winfo_width(), self.root.winfo_height()
                win_w, win_h = help_win.winfo_width(), help_win.winfo_height()
                x = root_x + (root_w // 2) - (win_w // 2)
                y = root_y + (root_h // 2) - (win_h // 2)
                help_win.geometry(f"+{x}+{y}")
            else: help_win.geometry(f"+100+100")
            help_win.wait_window()
        except tk.TclError:
            try:
                if self.root.winfo_exists(): messagebox.showerror(APP_TITLE, "لم يتمكن من فتح نافذة المساعدة.", parent=self.root)
            except tk.TclError: pass
        except Exception as e:
            try:
                if self.root.winfo_exists(): messagebox.showerror(APP_TITLE, f"حدث خطأ غير متوقع عند فتح المساعدة:\n{e}", parent=self.root)
            except tk.TclError: pass

    def add_context_menu(self, widget):
        menu = tk.Menu(widget, tearoff=0, font=self.menu_font_spec)
        is_text_scrolled = isinstance(widget, (tk.Text, scrolledtext.ScrolledText))
        is_listbox = isinstance(widget, tk.Listbox)
        is_editable = can_select_all = can_copy = False
        try:
            if is_text_scrolled:
                is_editable = widget.cget('state') == tk.NORMAL
                can_select_all = can_copy = True
            elif is_listbox: can_copy = True
        except tk.TclError: pass

        if is_editable:
            menu.add_command(label=CUT_TEXT, command=lambda w=widget: w.event_generate("<<Cut>>"), accelerator="Ctrl+X", state=tk.DISABLED)
            menu.add_command(label=COPY_TEXT, command=lambda w=widget: w.event_generate("<<Copy>>"), accelerator="Ctrl+C", state=tk.DISABLED)
            menu.add_command(label=PASTE_TEXT, command=lambda w=widget: w.event_generate("<<Paste>>"), accelerator="Ctrl+V", state=tk.NORMAL)
        elif can_copy:
            cmd = lambda w=widget: self.copy_listbox_selection(w) if is_listbox else self.copy_from_disabled_text(w)
            menu.add_command(label=COPY_TEXT, command=cmd, accelerator="Ctrl+C", state=tk.DISABLED)

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
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)) and widget.tag_ranges(tk.SEL): has_selection = True
            elif isinstance(widget, tk.Listbox) and widget.curselection(): has_selection = True
            
            is_editable = False
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                 try: is_editable = widget.cget('state') == tk.NORMAL
                 except tk.TclError: pass

            for item_label in [CUT_TEXT, COPY_TEXT, PASTE_TEXT, SELECT_ALL_TEXT]:
                try:
                    item_idx = menu.index(item_label)
                    if item_idx is not None:
                        new_state = tk.DISABLED
                        if item_label == CUT_TEXT: new_state = tk.NORMAL if has_selection and is_editable else tk.DISABLED
                        elif item_label == COPY_TEXT: new_state = tk.NORMAL if has_selection else tk.DISABLED
                        elif item_label == PASTE_TEXT: new_state = tk.NORMAL if is_editable else tk.DISABLED
                        elif item_label == SELECT_ALL_TEXT:
                             if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                                 try: new_state = tk.NORMAL if widget.get("1.0", "end-1c").strip() else tk.DISABLED
                                 except tk.TclError: pass
                             else: new_state = tk.DISABLED 
                        menu.entryconfig(item_idx, state=new_state)
                except tk.TclError: continue
            menu.tk_popup(event.x_root, event.y_root)
        except tk.TclError: pass

    def copy_from_disabled_text(self, widget):
        if not isinstance(widget, (tk.Text, scrolledtext.ScrolledText)): return
        try:
            if not widget.winfo_exists() or not self.root.winfo_exists(): return
            if widget.tag_ranges(tk.SEL):
                self.root.clipboard_clear()
                self.root.clipboard_append(widget.get(tk.SEL_FIRST, tk.SEL_LAST))
                self.root.update()
        except tk.TclError as e: print(f"Error copying from disabled text widget: {e}", file=sys.stderr)

    def copy_listbox_selection(self, listbox):
        try:
             if not listbox.winfo_exists() or not self.root.winfo_exists(): return
             sel_indices = listbox.curselection()
             if not sel_indices: return
             self.root.clipboard_clear()
             self.root.clipboard_append("\n".join([listbox.get(i) for i in sel_indices]))
             self.root.update()
        except tk.TclError as e: print(f"Error copying from listbox: {e}", file=sys.stderr)

    def select_all(self, widget):
        if not isinstance(widget, (tk.Text, scrolledtext.ScrolledText)) or not widget.winfo_exists(): return 'break'
        original_state = None
        try:
            if widget.cget('state') == tk.DISABLED:
                original_state = tk.DISABLED
                widget.config(state=tk.NORMAL)
            widget.tag_remove(tk.SEL, "1.0", tk.END)
            widget.tag_add(tk.SEL, "1.0", "end-1c") 
            widget.mark_set(tk.INSERT, "1.0"); widget.see(tk.INSERT)
        except tk.TclError: pass
        finally:
             try:
                 if original_state == tk.DISABLED and widget.winfo_exists(): widget.config(state=tk.DISABLED)
             except tk.TclError: pass
        return 'break'

    def on_closing(self):
        if not self.root.winfo_exists(): return
        try:
            if self.worker_thread and self.worker_thread.is_alive():
               if messagebox.askokcancel("خروج", "معالجة الملفات لا تزال قيد التشغيل. هل تريد الخروج على أي حال؟\n(قد لا تكتمل معالجة الملف الحالي)", parent=self.root):
                   self.root.destroy()
            else: self.root.destroy()
        except tk.TclError:
            try:
                if self.root.winfo_exists(): self.root.destroy()
            except tk.TclError: pass
        except Exception as e:
            print(f"Unexpected error during closing: {e}", file=sys.stderr)
            try:
                if self.root.winfo_exists(): self.root.destroy()
            except tk.TclError: pass

# --- Main Execution ---
if __name__ == "__main__":
    try:
        if sys.stdout.encoding is None or sys.stdout.encoding.lower() != 'utf-8': sys.stdout.reconfigure(encoding='utf-8')
        if sys.stderr.encoding is None or sys.stderr.encoding.lower() != 'utf-8': sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError: pass 
    except Exception as e: print(f"Note: Could not reconfigure stdout/stderr encoding to UTF-8: {e}", file=sys.stderr)

    try:
        if DND_SUPPORTED:
            root = TkinterDnD.Tk()
        else:
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

