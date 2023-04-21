import sys
import openpyxl as op
import pandas as pd

from PyQt5.QtWidgets import *
from PyQt5 import uic



form_class = uic.loadUiType("ui/test1.ui")[0]


# Qtwidgets의 QMainWindow, ui파일의 form_class 상속
class WindowClass(QDialog, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # UI Setup

        # fileSelect 버튼 클릭시 selectFunction 메서드 동작
        self.fileSelect.clicked.connect(self.selectFunction)


    # selectFunction 메서드 정의
    def selectFunction(self):
        # filePath 출력하는 부분 초기화
        self.filePath.clear()
        # comboBox 출력하는 부분 초기화
        # 선택한 엑셀 파일 경로를 받아옴 : 튜플 타입으로 받아오며 0번재 요소가 주소값 string이다.
        path = QFileDialog.getOpenFileName(self, 'Open File', '', 'All File(*);; xlsx File(*.xlsx)')
        # filePath에 현재 읽어온 엑셀 파일 경로를 입력한다.(절대경로)
        self.filePath.setText(path[0])

        # 위 절대 경로 활용해 openpyxl workbook 객체 생성
        wb = op.load_workbook(path[0])
        # 설정한 workbook의 시트리스트를 읽어온다.
        self.shtlist = wb.sheetnames
        print(self.shtlist)


    def bga_df_maker(self, df):
        pin_int = df.columns.to_list()      # len : 27
        pin_int.pop(0)
        pin_cha = df["Unnamed: 0"].to_list()# len : 27

        pin_number = []     # 핀넘버
        for c in pin_cha:
            for i in pin_int:
                i = str(i)
                pin_number.append(c+i)
        # print(df.loc[0])
        # print(df.loc[0].to_list())
        pin_name_all = []

        for i in range(0, len(df["1"])):    # len(pin_int) = 27
            temp_list = df.loc[i].to_list()
            temp_list.pop(0)
            pin_name_all.append(temp_list)

        pin_name = []
        for i in pin_name_all:
            pin_name.extend(i)

        mydf = pd.DataFrame({"pin number" : pin_number, "pin name" : pin_name})
        out_df = mydf.dropna(axis=0)

        return out_df


# GUI 출력 부분
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()