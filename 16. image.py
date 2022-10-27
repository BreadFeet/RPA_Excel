from openpyxl import Workbook
from openpyxl.drawing.image import Image
wb = Workbook()
ws = wb.active

img = Image("capture.png")    # 이미지 파일 불러오기

ws.add_image(img, "C3")   # C3에 이미지 삽입

wb.save("sample_image.xlsx")