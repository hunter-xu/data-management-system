#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
import socket
import xlrd
import xlwt
import time
from xlutils import copy
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import matplotlib.pyplot as plt
#######################################################################

def save_date(tem, num):
    localtime = time.localtime(time.time())
    if not os.path.isfile("date/"+str(localtime.tm_year) + ".xls"):
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet(str(localtime.tm_mon))
        worksheet.write(0, 0, label='num')
        worksheet.write(0, 1, label='card_num')
        worksheet.write(0, 2, label='tem')
        worksheet.write(0, 3, label='day')
        worksheet.write(0, 4, label='time')
        workbook.save("date/"+str(localtime.tm_year) + ".xls")
    while not os.access("date/"+str(localtime.tm_year) + ".xls", os.R_OK):
        time.sleep(1)
    file_name = "date/"+str(localtime.tm_year) + ".xls"
    xl = xlrd.open_workbook(file_name)
    table_names = xl.sheet_names()
    if str(localtime.tm_mon) in table_names:
        table = xl.sheet_by_name(str(localtime.tm_mon))
        nrows = table.nrows
        wbook = copy.copy(xl)
        table_num = 0
        for i in range(len(table_names)):
            if str(localtime.tm_mon) == table_names[i]:
                table_num = i
        w_table = wbook.get_sheet(table_num)
        w_table.write(nrows, 0, nrows)
        w_table.write(nrows, 1, num)
        w_table.write(nrows, 2, tem)
        w_table.write(nrows, 3, localtime.tm_mday)
        time_date = str(localtime.tm_hour) + ":" + str(localtime.tm_min) + ":" + str(localtime.tm_sec)
        w_table.write(nrows, 4, time_date)
        wbook.save(file_name)
    else:
        wbook = copy.copy(xl)
        worksheet = wbook.add_sheet(str(localtime.tm_mon))
        worksheet.write(0, 0, label='num')
        worksheet.write(0, 1, label='card_num')
        worksheet.write(0, 2, label='tem')
        worksheet.write(0, 3, label='day')
        worksheet.write(0, 4, label='time')
        worksheet.write(1, 0, 1)
        worksheet.write(1, 1, num)
        worksheet.write(1, 2, tem)
        worksheet.write(1, 3, localtime.tm_mday)
        time_date = str(localtime.tm_hour) + ":" + str(localtime.tm_min) + ":" + str(localtime.tm_sec)
        worksheet.write(1, 4, time_date)
        wbook.save(file_name)
def get_date(year1=0, month1=0, day1=0, year2=0, month2=0, day2=0):
    lis = [[], [], [],[]]
    while year1 <= year2:
        if year1 == year2:
            month3 = month2
        else:
            month3 = 12
        try:
            workbook = xlrd.open_workbook("date/"+str(year1) + ".xls")
        except:
            print("打开文件失败：" + str(year1))
        else:
            while month1 <= month3:
                try:
                    sheet = workbook.sheet_by_name(str(month1))
                    # print("打开表格月份："+str(month1))
                except:
                    print("月份数据表获取失败：" + str(month1))
                else:
                    if month1 == month3:
                        day3 = day2
                    else:
                        day3 = 31
                    value1 = sheet.col_values(1)
                    value2 = sheet.col_values(2)
                    value3 = sheet.col_values(3)
                    value4 = sheet.col_values(4)
                    count = 1
                    for i in value3[1:]:
                        if i <= day3 and i>=day1:
                            pass
                        else:
                            del value1[count]
                            del value2[count]
                            del value3[count]
                            count = count - 1
                        count = count + 1
                    lis[0].extend(value1[1:])
                    lis[1].extend(value2[1:])
                    lis[2].extend(value3[1:])
                    lis[3].extend(value4[1:])
                    day1 = 0
                finally:
                    month1 = month1 + 1
        finally:
            month1 = 1
            year1 = year1 + 1
    return lis
def Processing_data(lis):
    lis_card = []
    lis_averages = []
    lis_count = []
    for i in lis[0]:
        if i not in lis_card:
            lis_card.append(i)
            lis_averages.append(0)
            lis_count.append(0)
    count = 0
    for i in lis_card:
        count1 = 0
        for j in lis[0]:
            if int(i) == int(j):
                lis_averages[count] = (lis_averages[count] * lis_count[count] + lis[1][count1]) / (lis_count[count] + 1)
                lis_count[count] = lis_count[count] + 1
            count1 = count1 + 1
        count = count + 1
    li = [lis_card, lis_averages, lis_count]
    return li
def change_date(lis):
    change_lis=[]
    count=0
    for i in lis[0]:
        x=[lis[0][count],lis[1][count],lis[2][count]]
        change_lis.append(x)
        count=count+1
    return change_lis
def get_average(lis):
    i = 0
    sum_tem = 0
    count = 0
    while i < len(lis[1]):
        sum_tem = sum_tem + lis[1][i] * lis[2][i]
        count = count + lis[2][i]
        i = i + 1
    return sum_tem / count
def get_cr():
    x=xlrd.open_workbook("config/data.xls")
    sheet=x.sheet_by_name("Sheet1")
    x=sheet.nrows
    y=sheet.ncols
    value=sheet.row_values(0)
    values=[]
    for i in value:
        values.append(i)
        values.append("温度")
    return [x+1,y*2,values]
def start_a_poss(update_data_thread):
    '''多线程开始函数'''
    update_data_thread.date_sender.connect(ex.update_item_data)  # 链接信号
    update_data_thread.start()

########################################################################

class mainwindow(QWidget):
    '''主窗口类，实现程序的主要窗口'''
    def __init__(self):
        super().__init__()
        self.initUI()
    def initUI (self):
        self.s=Show_on_time()
        self.resize(720, 900)
        self.setFixedSize(720,900)
        self.center()
        self.setWindowTitle('温度数据管理系统')
        self.setWindowIcon(QIcon('img/logo.png'))
        self.grid = QGridLayout()
        self.grid.setSpacing(20)
        self.ip = QLabel('ip地址')
        self.grid.addWidget(self.ip, 0, 0)
        self.ipEdit = QLineEdit('10.10.100.254')
        self.grid.addWidget(self.ipEdit, 0, 1, 1, 2)
        self.port = QLabel('端口号')
        self.grid.addWidget(self.port, 2, 0)
        self.portEdit = QLineEdit('8899')
        self.grid.addWidget(self.portEdit, 2, 1, 1, 2)
        self.btn = QPushButton("连接")
        self.btn.clicked.connect(self.tcpconnect)
        self.grid.addWidget(self.btn, 0, 3, 1, 2)
        self.btn2 = QPushButton("关闭连接")
        self.btn2.clicked.connect(self.close_tcpconnect)
        self.grid.addWidget(self.btn2, 2, 3, 1, 2)
        self.btn2.setEnabled(False)
        # 右侧按钮
        self.daybtn = QPushButton("日报表")
        self.grid.addWidget(self.daybtn, 0, 6, 1, 2)
        self.daybtn.clicked.connect(self.day_date_show)
        self.mouthbtn = QPushButton("月报表")
        self.grid.addWidget(self.mouthbtn, 0, 8, 1, 2)
        self.mouthbtn.clicked.connect(self.mon_date_show)
        self.yearbtn = QPushButton("总报表")
        self.grid.addWidget(self.yearbtn, 2, 6, 1, 2)
        self.yearbtn.clicked.connect(self.sum_date_show)
        self.exitbtn = QPushButton("曲线数据")
        self.grid.addWidget(self.exitbtn, 2, 8, 1, 2)
        self.exitbtn.clicked.connect(self.show_img)
        self.show_class=MyWindow2()
        self.show_imgs = QpixmapDemo()
        self.x=get_cr()
        self.tableWidget = QTableWidget(100,5)
        self.tableWidget.setHorizontalHeaderLabels(["序号","槽号","温度","日期","时间"])
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.grid.addWidget(self.tableWidget,6, 0, 5, 10)
        self.setLayout(self.grid)
    def show_img(self):
            self.show_imgs.show()
    def update_item_data(self, data):
        """更新内容"""
        self.tableWidget.setItem(int(data[0]), 0, QTableWidgetItem(str(data[0])))
        self.tableWidget.setItem(int(data[0]), 1, QTableWidgetItem(str(data[1])))
        self.tableWidget.setItem(int(data[0]), 2, QTableWidgetItem(str(data[2])))
        r=str(data[3].tm_year)+"年"+str(data[3].tm_mon)+"月"+str(data[3].tm_mday)+"日"
        self.tableWidget.setItem(int(data[0]), 3, QTableWidgetItem(r))
        t = str(data[3].tm_hour) + ":" + str(data[3].tm_min) + ":" + str(data[3].tm_sec)
        self.tableWidget.setItem(int(data[0]), 4, QTableWidgetItem(t))
        if data[0]==100:
            for j in range(5):
                for h in range(100):
                    self.tableWidget.setItem(h,j, QTableWidgetItem(" "))
    def day_date_show(self):
        dialog = day_date_show()
        res = dialog.exec_()
        date = dialog.datetime.date()
        if res==1:
            lis1=get_date(date.year(),date.month(),date.day(),date.year(),date.month(),date.day())
            lis2=Processing_data(lis1)
            lis3=change_date(lis2)
            self.show_class.set_show_date(lis3)
            self.show_class.show()
        else:
            pass
    def mon_date_show(self):
        dialog = month_date_show()
        res = dialog.exec_()
        date = dialog.datetime.date()
        if res==1:
            lis1 = get_date(date.year(), date.month(), 0, date.year(), date.month(), 30)
            lis2 = Processing_data(lis1)
            lis3 = change_date(lis2)
            self.show_class.set_show_date(lis3)
            self.show_class.show()
        else:
            pass
    def sum_date_show(self):
        begin_dialog=date_show()
        begin_res=begin_dialog.exec_()
        begin_date=begin_dialog.datetime.date()
        begin_time=begin_dialog.datetime.time()
        if begin_res==1:
            end_dialog = date_show(text="请选择结束日期")
            end_res = end_dialog.exec_()
            end_date = end_dialog.datetime.date()
            end_time = end_dialog.datetime.time()
            if end_res==1:
                lis1 = get_date(begin_date.year(), begin_date.month(), begin_date.day(), end_date.year(),
                                end_date.month(), end_date.day())
                lis2 = Processing_data(lis1)
                lis3 = change_date(lis2)
                self.show_class.set_show_date(lis3)
                self.show_class.show()
            else:
                pass
        else :
            pass
    def show_base(self,lis):
        x=xlrd.open_workbook("config/data.xls")
        sheet=x.sheet_by_name("Sheet1")
        row=sheet.row_values(0)
        lie=[]
        for i in range(len(row)):
            lie.append(sheet.col_values(i))
        print(lie)
        for i in range(len(row)):
            count=0
            for j in lis:
                if str(j[0]) == str(lie[i][1]):
                    self.tableWidget.setItem(count, 2*i, QTableWidgetItem(str(j[0])))
                    self.tableWidget.setItem(count, 1 + 2 * i, QTableWidgetItem(str(j[1])))
                    count=count+1
    def tcpconnect(self):
        self.btn.setEnabled(False)
        ip=self.ipEdit.text()
        port=self.portEdit.text()
        self.btn2.setEnabled(True)
        self.s.set_up(ip,int(port))
        start_a_poss(self.s)
    def close_tcpconnect(self):
        try:
            if self.s.isRunning():
                self.s.close()
        except:
            pass
        finally:
            self.btn.setEnabled(True)
            self.btn2.setEnabled(False)
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    def closeEvent(self, event):
        reply = QMessageBox.question(self, '提示', "退出将不再收集传感数据，是否退出？", QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
            try:
                while self.p1.is_alive():
                    self.p1.kill()
            except:
                pass
        else:
            event.ignore()
class day_date_show(QDialog):
    def __init__(self, parent=None):
        super(day_date_show, self).__init__(parent)
        self.setWindowIcon(QIcon('img/logo.png'))
        self.setWindowTitle('时间')
        layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.datetime = QDateTimeEdit(self)
        self.datetime.setCalendarPopup(True)
        self.datetime.setDateTime(QDateTime.currentDateTime())
        self.label.setText("请选择要查看的日期")
        layout.addWidget(self.label)
        layout.addWidget(self.datetime)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    def ssss(self):
        self.accept()
class month_date_show(QDialog):
    def __init__(self, parent=None):
        super(month_date_show, self).__init__(parent)
        self.setWindowIcon(QIcon('img/logo.png'))
        layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.datetime = QDateTimeEdit(self)
        self.datetime.setCalendarPopup(True)
        self.datetime.setDateTime(QDateTime.currentDateTime())
        self.label.setText("请选择要查看的月份")
        layout.addWidget(self.label)
        layout.addWidget(self.datetime)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)  # 点击ok，隐士存在该方法
        buttons.rejected.connect(self.reject)  # 点击cancel，该方法默认隐士存在
        layout.addWidget(buttons)
        # 该方法在父类方法中调用，直接打开了子窗体，返回值则用于向父窗体数据的传递
class date_show(QDialog):
    def __init__(self, parent=None,text="请选择起始日期"):
        super(date_show, self).__init__(parent)
        self.setWindowIcon(QIcon('img/logo.png'))
        layout = QVBoxLayout(self)
        self.label = QLabel(self)
        self.datetime = QDateTimeEdit(self)
        self.datetime.setCalendarPopup(True)
        self.datetime.setDateTime(QDateTime.currentDateTime())
        self.label.setText(text)
        layout.addWidget(self.label)
        layout.addWidget(self.datetime)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)  # 点击ok，隐士存在该方法
        buttons.rejected.connect(self.reject)  # 点击cancel，该方法默认隐士存在
        layout.addWidget(buttons)
        # 该方法在父类方法中调用，直接打开了子窗体，返回值则用于向父窗体数据的传递
class MyWindow2(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('img/logo.png'))
        self.setWindowTitle('报表')
        self.resize(900, 900)
        self.grid = QGridLayout()
        self.grid.setSpacing(20)
        self.x=get_cr()
        self.tableWidget = QTableWidget(self.x[0],self.x[1])
        self.tableWidget.setHorizontalHeaderLabels(self.x[2])
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.verticalHeader().setVisible(False)
    def set_show_date(self,lis):
        x = xlrd.open_workbook("config/data.xls")
        sheet = x.sheet_by_name("Sheet1")
        row = sheet.row_values(0)
        lie = []
        for i in range(self.x[0]):
            for j in range(self.x[1]):
                self.tableWidget.setItem(i,j, QTableWidgetItem(" "))
        for i in range(len(row)):
            lie.append(sheet.col_values(i))
        for i in range(len(row)):
            count = 0
            for j in lis:
                if int(j[0]) == (lie[i][1]):
                    self.tableWidget.setItem(count, 2*i, QTableWidgetItem(j[0]))
                    self.tableWidget.setItem(count, 1 + 2 * i, QTableWidgetItem(str(j[1])))
                    count = count + 1
        self.grid.addWidget(self.tableWidget, 0, 0, 5, 10)
        self.setLayout(self.grid)
class Show_on_time(QThread):
    '''实时发送并保存接收到的数据'''
    flag=0
    date_sender = pyqtSignal(list)
    def set_up(self,ip='10.10.100.254',port=8899):
        self.flag=0
        self.i=0
        self.ip=ip
        self.port=port
        self.addr=(ip,port)
        self.tcp_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    def close(self):
        self.flag=1
    def run(self):
        try:
            self.tcp_socket.connect(self.addr)
        except ZeroDivisionError:
            print(ZeroDivisionError)
        else:
            while True:
                if self.flag==1:
                    break
                num = ""
                data = self.tcp_socket.recv(1024)
                tem16 = data[1:3].hex()
                tem10 = int(tem16, 16) / 10
                for i in [1, 2, 3, 4]:
                    num16 = data[1 + 2 * i:3 + 2 * i].hex()
                    num10 = int(num16, 16) % 1000
                    num = num + str(num10)
                save_date(tem10, num)
                localtime = time.localtime(time.time())
                self.date_sender.emit([self.i,num,tem10,localtime])
                self.tcp_socket.recv(1024)
                self.i=self.i+1
                if self.i==100:
                    i=0
        finally:
            print("close")
class QpixmapDemo(QWidget):
    def __init__(self,parent=None):
        super(QpixmapDemo, self).__init__(parent)
        self.setWindowIcon(QIcon('img/logo.png'))
        self.setWindowTitle('折线图')
        self.grid = QGridLayout()
        self.grid.setSpacing(10)
        self.id = QLabel('卡号：')
        self.grid.addWidget(self.id, 0, 0)
        self.idEdit = QLineEdit('1444924113')
        self.grid.addWidget(self.idEdit, 1, 0, 1, 2)
        self.begin_time = QLabel('起始时间：')
        self.grid.addWidget(self.begin_time, 2, 0)
        self.begin_time_Edit = QLabel('暂无选择')
        self.grid.addWidget(self.begin_time_Edit, 3, 0, 1, 2)
        self.btn = QPushButton("选择时间")
        self.btn.clicked.connect(self.chose_begin_time)
        self.grid.addWidget(self.btn, 4, 0, 1, 2)
        self.end_time = QLabel('末时间：')
        self.grid.addWidget(self.end_time, 5, 0)
        self.end_time_Edit = QLabel('暂无选择')
        self.grid.addWidget(self.end_time_Edit, 6, 0, 1, 2)
        self.btn1 = QPushButton("选择时间")
        self.btn1.clicked.connect(self.chose_end_time)
        self.grid.addWidget(self.btn1, 7, 0, 1, 2)
        self.lab1 = QLabel()
        self.lab1.setPixmap(QPixmap('img//123.jpg'))
        self.grid.addWidget(self.lab1, 0, 3, 20, 10)
        self.btn2 = QPushButton("查看")
        self.btn2.clicked.connect(self.show_img)
        self.grid.addWidget(self.btn2, 8, 0, 1, 2)
        self.setLayout(self.grid)
    def chose_begin_time(self):
        dialog = day_date_show()
        res = dialog.exec_()
        date = dialog.datetime.date()
        if res==1:
            self.begin=date
            text=str(date.year())+"年"+str(date.month())+"月"+str(date.day())+"日"
            self.begin_time_Edit.setText(text)
        else :
            pass
    def chose_end_time(self):
        dialog = day_date_show()
        res = dialog.exec_()
        date = dialog.datetime.date()
        if res==1:
            self.end=date
            text=str(date.year())+"年"+str(date.month())+"月"+str(date.day())+"日"
            self.end_time_Edit.setText(text)
        else :
            pass
    def show_img(self):
        lis1=get_date(self.begin.year(),self.begin.month(),self.begin.day(),self.end.year(),self.end.month(),self.end.day())
        self.set_img(self.idEdit.text(),lis1)
    def set_img(self,card,lis):
        lis1=[[],[],[]]
        for i in range(len(lis[0])):
            if int(lis[0][i])==int(card):
                lis1[0].append(i)
                lis1[1].append(lis[1][i])
                lis1[2].append(str(lis[3][i]))
        #求最大最小平均
        sum=0
        max=0
        min=lis1[1][0]
        count=len(lis1[1])
        for i in lis1[1]:
            sum=sum+i
            if i > max:
                max=i
            if i < min :
                min = i
        pingjun=round(sum/count,2)
        print(pingjun)
        plt.figure(figsize=(15,7))
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['axes.unicode_minus'] = False
        plt.plot(lis1[2], lis1[1],marker='o')
        for a, b in zip(lis1[2], lis1[1]):
            plt.text(a, b, b, ha='center', va='bottom', fontsize=10)
        plt.xticks(rotation=90)
        plt.ylabel('温度')
        plt.text(lis1[2][0],pingjun+(max-min)/10, '□ 平均值：'+str(pingjun))
        plt.text(lis1[2][0], pingjun+(max-min)*2/10, '□ 最大值：' + str(max))
        plt.text(lis1[2][0], pingjun+(max-min)*3/10, '□ 最小值：' + str(min))
        plt.savefig("img//new.jpg")
        self.lab1.setPixmap(QPixmap('img//new.jpg'))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = mainwindow()
    ex.show()
    sys.exit(app.exec_())