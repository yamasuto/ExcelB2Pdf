import sys
import os
import win32gui
import win32com.client
import pyautogui as pg
import time
import ctypes
import glob
import csv

# 元帳に対する操作毎に待つ時間 
M_OPERATION_INTERVAL_SEC = 0.5
# 上に加えて更に待つ時間
M_FILTER_INTERVAL_SEC = 0.5

def foreground(hwnd, title):
    name = win32gui.GetWindowText(hwnd)
    if name.find(title) >= 0:
        # if win32gui.IsIconic(hwnd):
        #     win32gui.ShowWindow(hwnd,1) # SW_SHOWNORMAL
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        print(f"{title} was set foreground")
        return False
    return True

def get_元帳科目list(worksheet):
    l = []
    row = 12
    while True:
        cell = worksheet.cells(row, 9) # I12から
        if cell.Text == "" or cell.Text == "- ":
            break
        s = cell.Text.split()
        if s not in l:
            l.append(s)
        row = row + 1
    return l

def read_csv_file(file_path):
    """
    指定されたCSVファイルを読み込む関数。
    
    Args:
    - file_path (str): 読み込むCSVファイルのパス。
    
    Returns:
    - list: CSVファイルから読み込まれたデータのリスト。
    """
    data = []
    with open(file_path, 'r', newline='', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            data.append(row)
    return data

def write_list_to_csv(data, file_path):
    """
    リストをCSVファイルに書き込む関数。
    
    Args:
    - data (list of lists): CSVファイルに書き込むデータ。
    - file_path (str): 書き込むCSVファイルのパス。
    """
    with open(file_path, 'w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        csv_writer.writerows(data)

def export_元帳(worksheet, excel_file_name_without_ext, count, pdf_folder_path, items_csv_file_path):

    # ここでExcelをアクティブにして、キー押下できるようにする
    # This will always raise that exception if your handler returns 0. 
    try:
        win32gui.EnumWindows(foreground, excel_file_name_without_ext)
    except win32gui.error:
        None

    # シートを選択する
    worksheet.Activate() 
    time.sleep(M_OPERATION_INTERVAL_SEC)

    if items_csv_file_path and os.path.exists(items_csv_file_path):
        print(f"元帳の科目をファイルから読み込む {items_csv_file_path}")
        l = read_csv_file(items_csv_file_path)
    else:
        print("元帳の科目を列挙する...")
        l = get_元帳科目list(worksheet)
        write_list_to_csv(l, os.path.join(pdf_folder_path, f"items_{excel_file_name_without_ext}.csv"))
    print(l)

    # F9を取得する
    cell = worksheet.cells(9,6)
    # print(cell.GetAddress(RowAbsolute=False, ColumnAbsolute=False))
    # 操作後(一処理するごと)に500ms待つ
    pg.PAUSE=M_OPERATION_INTERVAL_SEC
    skipped_items = []
    files = 0
    for item in l:
        cell.Activate()
        time.sleep(M_OPERATION_INTERVAL_SEC)       
        pg.hotkey('alt','down')
        time.sleep(M_FILTER_INTERVAL_SEC)
        pg.press('tab')
        pg.press('tab')
        pg.press('tab')
        pg.press('tab')
        pg.typewrite(item[0])
        print(f"{item[0]}-{item[1]}")
        time.sleep(M_FILTER_INTERVAL_SEC)
        pg.hotkey('enter')
        time.sleep(M_FILTER_INTERVAL_SEC)

        pdf_file_name = f"{excel_file_name_without_ext}_{count}_{worksheet.Name}_{item[0]}{item[1]}.pdf"
        export_pdf(worksheet, os.path.join(pdf_folder_path,pdf_file_name))
        files = files + 1
    
    if len(skipped_items):
        print(f"The following items does not exist : {skipped_items}")
    
    return files

def export_pdf(worksheet, path):
    print(path)
    # xlTypePDF = 0
    worksheet.ExportAsFixedFormat(0, path)

def main(excel_file_path, pdf_folder_path, items_csv_file_path):
    excel_file_path = excel_file_path.strip()
    pdf_folder_path = pdf_folder_path.strip()
    if items_csv_file_path:
        items_csv_file_path = items_csv_file_path.strip()

    if not os.path.exists(excel_file_path):
        print(f"The Excel file does not exist. [{excel_file_path}]")
        return

    if not os.path.exists(pdf_folder_path):
        print(f"The Output folder does not exist. [{pdf_folder_path}]")
        return

    pdf_files = glob.glob(os.path.join(pdf_folder_path, "*.pdf"))
    if len(pdf_files) > 0:
        print(f"The Output folder contains {pdf_files} PDF files. [{pdf_folder_path}]")
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False #警告を表示しないようにする
    excel_file_name_without_ext = os.path.splitext(os.path.basename(excel_file_path))[0]

    count = 0
    fc = 0
    try:
        workbook = excel.Workbooks.Open(excel_file_path)
        for worksheet in workbook.Worksheets:
            if worksheet.Name in ["はじめに", "ⅱ", "台帳様式", "精算表", "クエリ",]:
                continue

            count += 1

            if worksheet.Name.startswith("元帳ⅰ"):
                fc = export_元帳(worksheet, excel_file_name_without_ext, count, pdf_folder_path, items_csv_file_path)
                continue

            pdf_file_name = f"{excel_file_name_without_ext}_{count}_{worksheet.Name}.pdf"
            export_pdf(worksheet, os.path.join(pdf_folder_path, pdf_file_name))

        workbook.Close()
        print(f"Exported {count} sheets to {count+fc-1} PDF files")
    finally:
        if excel is not None:
            excel.Quit()
            excel = None

        print("Finished.")

if __name__ == "__main__":
    # 以下使うため、インストールしてなければ実行する
    # pip install pywin32          // https://github.com/mhammond/pywin32
    # pip install pyautogui        // https://github.com/asweigart/pyautogui

	# 「元帳ⅰ」シートはすべて選択状態にしておく∵export_元帳でExcelのフィルターに文字列を入力する

    # 期待通り動かないときは、スクリプトを読み怪しそうなところにSleepを追加してみる

    if len(sys.argv) < 3:
        print("Usage: python ExcelB2Pdfs.py <excel_file_path> <pdf_folder_path> [<items_csv_file_path>]")
    else:
        if len(sys.argv) == 3:
            main(sys.argv[1], sys.argv[2], None)
        if len(sys.argv) > 3:
            main(sys.argv[1], sys.argv[2], sys.argv[3])
