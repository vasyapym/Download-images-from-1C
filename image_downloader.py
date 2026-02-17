#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# === EARLY DIAGNOSTIC ===
print('=' * 60)
print('  СКРИПТ ЗАПУЩЕН — проверяю зависимости...')
print('=' * 60)

import os
import sys
import re
import math
from urllib.parse import urlparse

try:
    import requests
except ImportError:
    print('\nОШИБКА: Библиотека requests не установлена!')
    print('Выполните в cmd:  pip install requests')
    input('\nНажмите Enter...')
    sys.exit(1)

try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    pass

HAVE_OPENPYXL = False
try:
    import openpyxl
    HAVE_OPENPYXL = True
except ImportError:
    pass

HAVE_PANDAS = False
try:
    import pandas as pds
    HAVE_PANDAS = True
except ImportError:
    pass

HAVE_XLRD = False
try:
    import xlrd
    HAVE_XLRD = True
except ImportError:
    pass

print('  requests  OK')
if HAVE_OPENPYXL:
    print('  openpyxl  OK')
if HAVE_PANDAS:
    print('  pandas    OK')
if HAVE_XLRD:
    print('  xlrd      OK')
print()


# ================================================================
#  TRANSLITERATION: Russian -> Latin
# ================================================================
TRANSLIT_MAP = {
    'а': 'a',  'б': 'b',  'в': 'v',  'г': 'g',  'д': 'd',
    'е': 'e',  'ё': 'yo', 'ж': 'zh', 'з': 'z',  'и': 'i',
    'й': 'y',  'к': 'k',  'л': 'l',  'м': 'm',  'н': 'n',
    'о': 'o',  'п': 'p',  'р': 'r',  'с': 's',  'т': 't',
    'у': 'u',  'ф': 'f',  'х': 'kh', 'ц': 'ts', 'ч': 'ch',
    'ш': 'sh', 'щ': 'shch', 'ъ': '',  'ы': 'y',  'ь': '',
    'э': 'e',  'ю': 'yu', 'я': 'ya',
    'А': 'A',  'Б': 'B',  'В': 'V',  'Г': 'G',  'Д': 'D',
    'Е': 'E',  'Ё': 'Yo', 'Ж': 'Zh', 'З': 'Z',  'И': 'I',
    'Й': 'Y',  'К': 'K',  'Л': 'L',  'М': 'M',  'Н': 'N',
    'О': 'O',  'П': 'P',  'Р': 'R',  'С': 'S',  'Т': 'T',
    'У': 'U',  'Ф': 'F',  'Х': 'Kh', 'Ц': 'Ts', 'Ч': 'Ch',
    'Ш': 'Sh', 'Щ': 'Shch', 'Ъ': '',  'Ы': 'Y',  'Ь': '',
    'Э': 'E',  'Ю': 'Yu', 'Я': 'Ya',
}


def transliterate(text):
    return ''.join(TRANSLIT_MAP.get(ch, ch) for ch in text)


def sanitize_filename(name, max_len=120):
    name = transliterate(name)
    name = re.sub(r'[^\w\-]', '_', name, flags=re.ASCII)
    name = re.sub(r'_+', '_', name)
    name = name.strip('_')
    if not name:
        return ''
    return name[:max_len]


# ================================================================
#  SAFE CELL READER (handles NaN, None, etc.)
# ================================================================
def cell_to_str(value):
    if value is None:
        return ''
    try:
        if isinstance(value, float) and math.isnan(value):
            return ''
    except (TypeError, ValueError):
        pass
    s = str(value).strip()
    if s.lower() in ('none', 'nan', 'nat', '<na>'):
        return ''
    return s


# ================================================================
#  HEADER TOOLS
# ================================================================
def normalize_header(h):
    if h is None:
        return ''
    s = str(h)
    s = re.sub(r'[\u00a0\u200b\u200c\u200d\u2060\ufeff]', ' ', s)
    return s.lower().strip()


def find_col(headers, keywords):
    for idx, h in enumerate(headers):
        h_norm = normalize_header(h)
        if h_norm and all(kw.lower() in h_norm for kw in keywords):
            return idx
    return None


def is_autonumbered(row):
    if not row:
        return False
    for i, val in enumerate(row):
        s = cell_to_str(val)
        if s != str(i):
            return False
    return True


def looks_like_header(row):
    text_count = 0
    for val in row:
        s = cell_to_str(val)
        if s and not s.replace('.', '').replace('-', '').isdigit() and not s.startswith('http'):
            text_count += 1
    return text_count >= 2


# ================================================================
#  HELPERS
# ================================================================
def find_xlsx_files(directory):
    result = []
    for f in os.listdir(directory):
        low = f.lower()
        if (low.endswith('.xlsx') or low.endswith('.xls')) and not f.startswith('~$'):
            result.append(os.path.join(directory, f))
    return result


def ext_from_url(url):
    path = urlparse(url).path
    ext = os.path.splitext(path)[1]
    ext = ext.split('?')[0].split('#')[0]
    return ext if ext else '.jpg'


def is_image_url(value):
    s = cell_to_str(value)
    if not s or not s.startswith('http'):
        return False
    lower = s.lower().split('?')[0].split('#')[0]
    return any(lower.endswith(e) for e in
               ('.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg', '.JPG'))


def download_file(url, filepath):
    hdrs = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/124.0.0.0 Safari/537.36'
        )
    }
    resp = requests.get(url, headers=hdrs, timeout=60, stream=True, verify=False)
    resp.raise_for_status()
    with open(filepath, 'wb') as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)


def unique_path(filepath):
    if not os.path.exists(filepath):
        return filepath
    base, ext = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(base + '_' + str(counter) + ext):
        counter += 1
    return base + '_' + str(counter) + ext


# <-- CHANGED: build_base now accepts artikul_col
def build_base(row, name_col, artikul_col, row_num):
    raw_name = ''
    if name_col is not None and name_col < len(row):
        raw_name = cell_to_str(row[name_col])

    raw_artikul = ''                                          # <-- NEW
    if artikul_col is not None and artikul_col < len(row):    # <-- NEW
        raw_artikul = cell_to_str(row[artikul_col])           # <-- NEW

    # If Артикул exists and is NOT already inside Название, prepend it
    if raw_artikul and raw_name:                              # <-- NEW
        if raw_artikul not in raw_name:                       # <-- NEW
            raw_name = raw_artikul + ' ' + raw_name           # <-- NEW
    elif raw_artikul and not raw_name:                        # <-- NEW
        raw_name = raw_artikul                                # <-- NEW

    base = sanitize_filename(raw_name) if raw_name else ''
    if not base:
        base = 'product_row_' + str(row_num)
    return base


# ================================================================
#  FALLBACK: scan every cell for image URLs
# ================================================================
# <-- CHANGED: added artikul_col parameter
def scan_all_cells_for_urls(rows, name_col, artikul_col, images_dir):
    print('')
    print('  Конкретные столбцы картинок не найдены.')
    print('  Ищу URL изображений во ВСЕХ ячейках...')
    print('')
    success = 0
    failed = 0

    for row_num, row in enumerate(rows[1:], start=2):
        base = build_base(row, name_col, artikul_col, row_num)  # <-- CHANGED
        img_index = 0
        for col_idx, cell in enumerate(row):
            if is_image_url(cell):
                url = cell_to_str(cell)
                ext = ext_from_url(url)
                filename = base + '_' + str(img_index) + ext
                filepath = unique_path(os.path.join(images_dir, filename))
                print('  Строка ' + str(row_num) + ', столбец ' +
                      str(col_idx + 1) + ': ' + os.path.basename(filepath) +
                      ' ... ', end='', flush=True)
                try:
                    download_file(url, filepath)
                    print('OK')
                    success += 1
                except Exception as e:
                    print('ОШИБКА: ' + str(e))
                    failed += 1
                img_index += 1

    return success, failed


# ================================================================
#  MAIN
# ================================================================
def main():
    print('  СКАЧИВАНИЕ ИЗОБРАЖЕНИЙ ИЗ КАТАЛОГА')
    print('  Пробелы заменяются на подчеркивания (_)')
    print('  Именование: Артикул_Название_0, Артикул_Название_1, ...')  # <-- CHANGED
    print('=' * 60)
    print()

    # --- Locate the file ---
    xlsx_path = None

    if len(sys.argv) > 1:
        xlsx_path = sys.argv[1]
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        print('Папка скрипта: ' + script_dir)
        found = find_xlsx_files(script_dir)
        print('Найдено файлов .xlsx/.xls: ' + str(len(found)))

        if found:
            if len(found) == 1:
                xlsx_path = found[0]
            else:
                print()
                for i, f in enumerate(found, 1):
                    print('  ' + str(i) + '. ' + os.path.basename(f))
                choice = input('Введите номер файла: ').strip()
                try:
                    xlsx_path = found[int(choice) - 1]
                except (ValueError, IndexError):
                    xlsx_path = found[0]

    if not xlsx_path:
        print('ОШИБКА: .xlsx/.xls файл не найден!')
        input('\nНажмите Enter...')
        sys.exit(1)

    print('Файл: ' + xlsx_path)

    if not os.path.isfile(xlsx_path):
        print('ОШИБКА: Файл не существует!')
        input('\nНажмите Enter...')
        sys.exit(1)

    # --- Output folder ---
    output_dir = os.path.join(
        os.path.dirname(os.path.abspath(xlsx_path)), 'Downloaded_Images')
    os.makedirs(output_dir, exist_ok=True)
    print('Папка: ' + output_dir)
    print()

    # --- Open workbook ---
    rows = None

    with open(xlsx_path, 'rb') as f:
        file_header = f.read(500).lower()

    is_html = (b'<html' in file_header or b'<table' in file_header
               or b'<!doctype' in file_header)

    if is_html:
        print('INFO: Файл является HTML-таблицей (сохранен как .xls).')
        if not HAVE_PANDAS:
            print('ОШИБКА: pip install pandas lxml')
            input('\nНажмите Enter...')
            sys.exit(1)
        dfs = pds.read_html(xlsx_path, header=None)
        if not dfs:
            print('ОШИБКА: Не найдено таблиц в HTML!')
            input('\nНажмите Enter...')
            sys.exit(1)
        df = dfs[0]
        rows = df.values.tolist()
        print('  Строк в таблице: ' + str(len(rows)))

    elif xlsx_path.lower().endswith('.xls') and not xlsx_path.lower().endswith('.xlsx'):
        if HAVE_XLRD:
            print('INFO: Читаю .xls через xlrd...')
            xls_wb = xlrd.open_workbook(xlsx_path)
            xls_ws = xls_wb.sheet_by_index(0)
            rows = [xls_ws.row_values(r) for r in range(xls_ws.nrows)]
        elif HAVE_PANDAS:
            print('INFO: Читаю .xls через pandas...')
            df = pds.read_excel(xlsx_path, header=None)
            rows = df.values.tolist()
        else:
            print('ОШИБКА: pip install xlrd')
            input('\nНажмите Enter...')
            sys.exit(1)

    else:
        if HAVE_OPENPYXL:
            print('INFO: Читаю .xlsx через openpyxl...')
            wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
            ws = wb.active
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        elif HAVE_PANDAS:
            print('INFO: Читаю .xlsx через pandas...')
            df = pds.read_excel(xlsx_path, header=None)
            rows = df.values.tolist()
        else:
            print('ОШИБКА: pip install openpyxl')
            input('\nНажмите Enter...')
            sys.exit(1)

    if not rows or len(rows) < 2:
        print('ОШИБКА: Файл пуст или содержит только заголовок!')
        input('\nНажмите Enter...')
        sys.exit(1)

    # ------------------------------------------------------------------
    #  Detect if first row is pandas auto-numbered (0,1,2,...)
    # ------------------------------------------------------------------
    if is_autonumbered(rows[0]) and len(rows) > 2 and looks_like_header(rows[1]):
        print('  INFO: Обнаружена авто-нумерация pandas — пропускаю её.')
        rows = rows[1:]

    # --- Headers ---
    headers = [cell_to_str(h) for h in rows[0]]

    print()
    print('Столбцов: ' + str(len(headers)))
    print('Строк данных: ' + str(len(rows) - 1))
    print()
    print('Заголовки (для отладки):')
    for i, h in enumerate(headers):
        print('  [' + str(i) + '] = "' + h + '"')
    print()

    # --- Detect NAME column ---
    name_col = find_col(headers, ['название'])
    if name_col is not None:
        print('  Столбец "Название" найден: [' + str(name_col) + '] = "' + headers[name_col] + '"')
    else:
        name_col = find_col(headers, ['наименование'])
        if name_col is not None:
            print('  Столбец "Наименование" найден: [' + str(name_col) + '] = "' + headers[name_col] + '"')
        else:
            name_col = find_col(headers, ['name'])
            if name_col is not None:
                print('  Столбец "Name" найден: [' + str(name_col) + '] = "' + headers[name_col] + '"')
            else:
                h0 = normalize_header(headers[0]) if headers else ''
                if h0 in ('id', 'ид', 'код', 'code', '#', '0', '') and len(headers) > 1:
                    name_col = 1
                    print('  ВНИМАНИЕ: "Название" не найдено!')
                    print('  Столбец [0] = "' + headers[0] + '" похож на ID.')
                    print('  Использую столбец [1] = "' + headers[1] + '" для имён.')
                else:
                    name_col = 0
                    print('  ВНИМАНИЕ: "Название" не найдено, использую столбец [0]')

    # --- Detect ARTIKUL column ---                                       # <-- NEW BLOCK
    artikul_col = find_col(headers, ['артикул'])                          # <-- NEW
    if artikul_col is not None:                                           # <-- NEW
        print('  Столбец "Артикул" найден: [' + str(artikul_col) + '] = "' + headers[artikul_col] + '"')  # <-- NEW
    else:                                                                 # <-- NEW
        print('  Столбец "Артикул": НЕ НАЙДЕН (файлы будут именоваться только по Названию)')  # <-- NEW

    # --- Detect IMAGE columns ---
    preview_col = find_col(headers, ['картинка', 'анонс'])
    detail_col = find_col(headers, ['детальная', 'картинка'])

    if preview_col is None:
        preview_col = find_col(headers, ['анонс'])
    if detail_col is None:
        detail_col = find_col(headers, ['детальная'])

    if preview_col is None and detail_col is None:
        for idx, h in enumerate(headers):
            h_norm = normalize_header(h)
            if 'картинк' in h_norm or 'image' in h_norm or 'фото' in h_norm:
                if detail_col is None:
                    detail_col = idx

    print()
    print('  Название:            столбец [' + str(name_col) + '] = "' + headers[name_col] + '"')

    if artikul_col is not None:                                           # <-- NEW
        print('  Артикул:             столбец [' + str(artikul_col) + '] = "' + headers[artikul_col] + '"')  # <-- NEW
    else:                                                                 # <-- NEW
        print('  Артикул:             НЕ НАЙДЕН')                         # <-- NEW

    if preview_col is not None:
        print('  Картинка для анонса: столбец [' + str(preview_col) + '] = "' + headers[preview_col] + '"')
    else:
        print('  Картинка для анонса: НЕ НАЙДЕН')

    if detail_col is not None:
        print('  Детальная картинка:  столбец [' + str(detail_col) + '] = "' + headers[detail_col] + '"')
    else:
        print('  Детальная картинка:  НЕ НАЙДЕН')

    print()

    # --- Show first few data rows for verification (no prompt) ---       # <-- CHANGED BLOCK
    print('  [ПРОВЕРКА] Первые 3 строки данных:')
    for test_i in range(1, min(4, len(rows))):
        test_row = rows[test_i]
        test_raw = cell_to_str(test_row[name_col]) if name_col < len(test_row) else '???'
        test_artikul = cell_to_str(test_row[artikul_col]) if artikul_col is not None and artikul_col < len(test_row) else ''  # <-- NEW
        test_base = build_base(test_row, name_col, artikul_col, test_i)   # <-- CHANGED
        print('    Строка ' + str(test_i + 1) + ': Название="' + test_raw + '", Артикул="' + test_artikul + '"')  # <-- CHANGED
        print('             -> файл: "' + test_base + '_0.jpg"')
    print()

    # --- Fallback if no image columns ---
    if preview_col is None and detail_col is None:
        success, failed = scan_all_cells_for_urls(rows, name_col, artikul_col, output_dir)  # <-- CHANGED
        print()
        print('=' * 50)
        print('  Скачано: ' + str(success) + '   Ошибок: ' + str(failed))
        print('  Папка:   ' + output_dir)
        print('=' * 50)
        input('\nНажмите Enter...')
        return

    # --- Process each data row ---
    success = 0
    failed = 0
    no_url = 0

    for row_num, row in enumerate(rows[1:], start=2):
        base = build_base(row, name_col, artikul_col, row_num)  # <-- CHANGED

        row_had_url = False
        img_index = 0

        # PREVIEW ("Картинка для анонса") -> always _0
        if (preview_col is not None
                and preview_col < len(row)
                and is_image_url(row[preview_col])):
            row_had_url = True
            url = cell_to_str(row[preview_col])
            ext = ext_from_url(url)
            filename = base + '_' + str(img_index) + ext
            filepath = unique_path(os.path.join(output_dir, filename))
            img_index += 1

            print('  [' + str(row_num) + '] (анонс)     ' +
                  os.path.basename(filepath) + '  ... ', end='', flush=True)
            try:
                download_file(url, filepath)
                print('OK')
                success += 1
            except Exception as e:
                print('ОШИБКА: ' + str(e))
                failed += 1

        # DETAIL ("Детальная картинка") -> _1 or _0
        if (detail_col is not None
                and detail_col < len(row)
                and is_image_url(row[detail_col])):
            row_had_url = True
            url = cell_to_str(row[detail_col])
            ext = ext_from_url(url)
            filename = base + '_' + str(img_index) + ext
            filepath = unique_path(os.path.join(output_dir, filename))
            img_index += 1

            print('  [' + str(row_num) + '] (детальная) ' +
                  os.path.basename(filepath) + '  ... ', end='', flush=True)
            try:
                download_file(url, filepath)
                print('OK')
                success += 1
            except Exception as e:
                print('ОШИБКА: ' + str(e))
                failed += 1

        if not row_had_url:
            no_url += 1

    # --- Summary ---
    print()
    print('=' * 60)
    print('  Успешно скачано:      ' + str(success))
    print('  Ошибок:               ' + str(failed))
    print('  Строк без картинок:   ' + str(no_url))
    print('  Файлы сохранены в:    ' + output_dir)
    print('=' * 60)
    input('\nНажмите Enter для выхода...')


# ================================================================
if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print()
        print('!!! НЕОБРАБОТАННАЯ ОШИБКА !!!')
        print(str(e))
        import traceback
        traceback.print_exc()
        input('\nНажмите Enter...')
