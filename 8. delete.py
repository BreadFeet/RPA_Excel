from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# 6번 학생 성적 삭제하기
# ws.delete_rows(7)

# 7번 학생부터 3명 삭제하기
# ws.delete_rows(8, 3)

# B컬럼 삭제
# ws.delete_cols(2)

# B컬럼에서 2열 삭제
ws.delete_cols(2, 2)

wb.save("sample_delete.xlsx")