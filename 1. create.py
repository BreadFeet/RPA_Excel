# pip install openpyxl 설치from open
from openpyxl import Workbook
wb = Workbook()     # 새 워크북 생성-빈 Excel 통합문서
ws = wb.active      # 현재 활성화된 기본 sheet를 가져옴
ws.title = "SoriTable"      # Sheet 이름을 변경
wb.save("sample.xlsx")      # 파일이름 지정하여 저장
wb.close()             
# 실행하면 메뉴에 sample 파일 생성됨


