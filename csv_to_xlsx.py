import csv, openpyxl, sys, os

def main():
    if len(sys.argv) < 3:
        print("用法: python csv_to_xlsx.py <來源.csv> <輸出.xlsx> [編碼]")
        sys.exit(1)

    src = sys.argv[1]
    dst = sys.argv[2]
    encodings = [sys.argv[3]] if len(sys.argv) > 3 else ['cp950', 'big5', 'utf-8-sig', 'utf-8']

    if not os.path.exists(src):
        print(f"錯誤: 找不到檔案: {src}")
        sys.exit(1)

    content = None
    used_enc = None
    for enc in encodings:
        try:
            with open(src, 'r', encoding=enc) as f:
                content = f.read()
            used_enc = enc
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if content is None:
        print(f"錯誤: 無法辨識檔案編碼，嘗試過: {encodings}")
        sys.exit(1)

    wb = openpyxl.Workbook()
    ws = wb.active
    reader = csv.reader(content.splitlines())
    for row in reader:
        ws.append(row)

    wb.save(dst)
    print(f"轉換完成（編碼: {used_enc}）: {os.path.basename(src)} -> {os.path.basename(dst)} ({ws.max_row} rows x {ws.max_column} cols)")

if __name__ == '__main__':
    main()
