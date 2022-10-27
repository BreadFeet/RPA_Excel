from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# 8행에 새로운 열 추가
# ws.insert_rows(8)

# 8행에 5줄 추가
# ws.insert_rows(8, 5)

# B컬럼에 새로운 열 추가
# ws.insert_cols(2)

# B컬럼에 새로운 3열 추가
ws.insert_cols(2, 3)

wb.save("sample_insert_row.xlsx")   # 원본 날아가지 않게 따로 저장