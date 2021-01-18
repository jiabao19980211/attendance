import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
import pandas as pd
from datetime import datetime, timedelta
import datetime
from datetime import datetime
import win32api

bgs_config = pd.read_excel("config.xlsx", 'Sheet3', header=4, engine='openpyxl')
cx_config = pd.read_excel("config.xlsx", 'Sheet2', header=4, engine='openpyxl')


class MainGUI(QtWidgets.QMainWindow):

    # 设置用户界面
    def __init__(self):
        super().__init__()  # 调用父类中__init__的函数
        self.setWindowTitle("考勤统计格式转换程序(要求配置文件config.xlsx与该.exe文件在同一目录下)")  # 设置窗口标题
        self.resize(800, 400)  # 设置窗口大小，单位为像素
        self.main_widget = QtWidgets.QWidget()  # 创建界面
        self.main_widget_layout = QtWidgets.QGridLayout()  # 选取布局为网格布局(多行多列)
        self.main_widget.setLayout(self.main_widget_layout)  # 设置布局
        # 设置部件
        self.input = QtWidgets.QLineEdit("请确保输入文件路径和名字正确且文件存在。选择文件后自动替换该语句")  # 文本编辑框1
        self.btn_Choose_Input_File = QtWidgets.QPushButton("选择文件(仅支持xlsx文件)")  # 按钮1及名称
        self.output = QtWidgets.QLineEdit("请确保输出文件路径正确且存在。选择路径后自动替换该语句")  # 文本编辑框2
        self.btn_Choose_Onput_Directory = QtWidgets.QPushButton("选择输出文件夹")  # 按钮2及名称
        self.editBox_Information_Display_Area = QtWidgets.QListWidget()  # 列表控件
        self.btn_Execution = QtWidgets.QPushButton("执行转换")  # 按钮3及名称
        # 部件的位置设置
        self.main_widget_layout.addWidget(self.input, 0, 0, 1, 2)  # 文本编辑框1放在第1行1列，占用1行2列
        self.main_widget_layout.addWidget(self.btn_Choose_Input_File, 0, 2, 1, 1)  # 按钮1放在第1行第3列，占用1行1列
        self.main_widget_layout.addWidget(self.output, 1, 0, 1, 2)
        self.main_widget_layout.addWidget(self.btn_Choose_Onput_Directory, 1, 2, 1, 1)
        self.main_widget_layout.addWidget(self.btn_Execution, 2, 2, 1, 1)
        self.main_widget_layout.addWidget(self.editBox_Information_Display_Area, 3, 0, 3, 3)

        self.setCentralWidget(self.main_widget)  # 设置QMainWindow的中心窗口

        self.btn_Choose_Input_File.clicked.connect(self.Fun_Choice_Input_File)  # 将"选择输入文件夹"按钮绑定Choice_dir_input函数
        self.btn_Choose_Onput_Directory.clicked.connect(
            self.Fun_Choice_Output_Directory)  # 将"选择输出文件夹"按钮绑定Fun_Choice_Output_Directory函数
        self.btn_Execution.clicked.connect(self.Fun_On_BtnClick_Execuation)  # “执行汇总”按钮绑定Fun_On_BtnClick_Execuation函数
        self.editBox_Information_Display_Area.itemClicked.connect(self.Fun_show_select_item)

    def Fun_show_select_item(self, item):
        if "不合法" in item.text():

            QMessageBox.information(self, "ListWidget", "请根据提示重新输入内容●●●●●●" + "系统提示: " + item.text())
        elif "程序" in item.text():
            QMessageBox.information(self, "ListWidget", "这是一条系统消息●●●●●●" + "系统提示: " + item.text().replace("-", ""))
        elif "输入" in item.text():  # 打开该次程序执行中使用的输入文件
            win32api.ShellExecute(0, 'open', item.text().replace("输入地址为：", ""), '', '', 1)
        elif "输出" in item.text():  # 打开该次程序执行中使用的输出文件夹
            os.startfile(item.text().replace("输出地址为：", ""))
        elif "文件1为：" in item.text():  # 打开这一次执行程序所产生的输出文件1
            win32api.ShellExecute(0, 'open', self.editBox_Information_Display_Area.item(
                self.editBox_Information_Display_Area.currentRow() - 1).text().replace("输出地址为：", "") + "/" + item.text().replace("文件1为：", ""), '', '', 1)
        elif "文件2为：" in item.text():  # 打开这一次执行程序所产生的输出文件2
            win32api.ShellExecute(0, 'open', self.editBox_Information_Display_Area.item(
                self.editBox_Information_Display_Area.currentRow() - 2).text().replace("输出地址为：","") + "/" + item.text().replace("文件2为：", ""), '', '', 1)
        else:
            QMessageBox.information(self, "ListWidget", "我是一条分割线●●●●●●" + "系统提示: " + item.text())
        print(item.text())

    def Fun_Choice_Input_File(self):
        p_str_InputFileName, p_str_InputFileType = QFileDialog.getOpenFileName(self, "选取文件", "D:\\", "*.xlsx")
        self.input.setText(p_str_InputFileName)

    def Fun_Choice_Output_Directory(self):
        p_str_OutputFilePath = QtWidgets.QFileDialog.getExistingDirectory(self, "请选择文件夹路径", "D:\\")
        self.output.setText(p_str_OutputFilePath)

    int_record_times = 0

    def Fun_On_BtnClick_Execuation(self):
        str_InputFile_Url = self.input.text()
        str_Output_Dir = self.output.text()
        if os.path.isfile(str_InputFile_Url) and os.path.isdir(str_Output_Dir):  # 判断输入内容是否是合法的
            strInputFileName = os.path.basename(str_InputFile_Url)[0:-5]
            data_stream_InputFile = pd.read_excel(str_InputFile_Url, engine='openpyxl')
            data_stream_InputFile['len'] = data_stream_InputFile.dropna(subset=['时间'])['时间'].str.split(' ').apply(len)
            data_stream_InputFile2 = data_stream_InputFile[data_stream_InputFile['len'] >= 1]
            bgs_data_stream_InputFile = data_stream_InputFile2[data_stream_InputFile2['len'] < 4]
            cx_data_stream_InputFile = data_stream_InputFile2[data_stream_InputFile2['len'] >= 4]
            bgs_data_stream_InputFile['状态'], bgs_data_stream_InputFile['平时上班'], bgs_data_stream_InputFile['平时加班'], \
            bgs_data_stream_InputFile['迟到'] = zip(*bgs_data_stream_InputFile['时间'].apply(bgs))
            cx_data_stream_InputFile['状态'], cx_data_stream_InputFile['平时上班'], cx_data_stream_InputFile['平时加班'], \
            cx_data_stream_InputFile['迟到'] = zip(*cx_data_stream_InputFile['时间'].apply(cx))
            bgs_data_stream_InputFile['平时上班'] = bgs_data_stream_InputFile['平时上班'] / 3600
            bgs_data_stream_InputFile['平时加班'] = bgs_data_stream_InputFile['平时加班'] / 3600
            bgs_data_stream_InputFile['迟到'] = bgs_data_stream_InputFile['迟到'] / 60
            cx_data_stream_InputFile['平时上班'] = cx_data_stream_InputFile['平时上班'] / 3600
            cx_data_stream_InputFile['平时加班'] = cx_data_stream_InputFile['平时加班'] / 3600
            cx_data_stream_InputFile['迟到'] = cx_data_stream_InputFile['迟到'] / 60

            time = str(datetime.now().strftime("%Y%m%d%H%M%S%f"))  # 创建一个datetime类

            str_output_file_path = str_Output_Dir + "/" + strInputFileName + time
            bgs_data_stream_InputFile.to_excel(str_output_file_path + "办公室人员统计.xlsx", 'w+')
            cx_data_stream_InputFile.to_excel(str_output_file_path + "产线人员统计.xlsx", 'w+')

            MainGUI.int_record_times = MainGUI.int_record_times + 1

            self.editBox_Information_Display_Area.insertItem(0,
                                                             "---------------------------------------------------------------------------------------")
            self.editBox_Information_Display_Area.insertItem(0, "文件2为：" + strInputFileName + time + "产线人员统计.xlsx")
            self.editBox_Information_Display_Area.insertItem(0, "文件1为：" + strInputFileName + time + "办公室人员统计.xlsx")
            self.editBox_Information_Display_Area.insertItem(0, "输出地址为：" + str_Output_Dir)
            self.editBox_Information_Display_Area.insertItem(0, "输入地址为：" + str_InputFile_Url)
            self.editBox_Information_Display_Area.insertItem(0, "程序执行完成，并且文件生成成功！！")
            self.editBox_Information_Display_Area.insertItem(0, "---------------程序第" + str(
                MainGUI.int_record_times) + "次执行----------------")

        else:  # 输入文件或者路径不合法的处理方法
            if os.path.isfile(str_InputFile_Url) != True:
                self.editBox_Information_Display_Area.insertItem(0, "输入文件路径不合法，或者文件不存在，请重新输入！！！")
            elif os.path.isdir(str_Output_Dir) != True:
                self.editBox_Information_Display_Area.insertItem(0, "输出文件路径不合法请重新输入！！！")
            else:
                self.editBox_Information_Display_Area.insertItem(0, "出现问题了，清确认输入")


def cat(lis_s, cats, how="()"):
    # 列表中的字符串转换为时间类型
    lis = []
    for ts in lis_s:
        lis.append(datetime.strptime(ts, '%H:%M'))

    # 列表中的字符串转换为时间类型
    cats_list = []
    for ts in cats:
        cats_list.append(datetime.strptime(ts, '%H:%M'))

    cat_result = [[] for _ in range(len(cats_list) - 1)]

    for t in lis:
        for i in range(len(cats_list)):
            if how == "(]":
                if cats_list[i] < t <= cats_list[i + 1]:
                    cat_result[i].append(t)
            elif how == '[)':
                if cats_list[i] <= t < cats_list[i + 1]:
                    cat_result[i].append(t)
            elif how == '()':
                if cats_list[i] < t < cats_list[i + 1]:
                    cat_result[i].append(t)
            elif how == '[]':
                if cats_list[i] <= t <= cats_list[i + 1]:
                    cat_result[i].append(t)
            else:
                raise ValueError('how 参数不正确')

    return cat_result


def bgs(s):
    a, b, c, d, e, f = cat(s.split(' '), '0:00 8:51 12:20 13:41 18:00 18:31 23:59'.split(' '), how='[)')
    al, bl, cl, dl, el, fl = [len(i) for i in [a, b, c, d, e, f]]
    ab, bb, cb, db, eb, fb = [0 if len(i) == 0 else 1 for i in [a, b, c, d, e, f]]
    v = bgs_config[
        (bgs_config['a'] == ab) & (bgs_config['b'] == bb) & (bgs_config['c'] == cb) & (bgs_config['d'] == db) & (
                bgs_config['e'] == eb) & (bgs_config['f'] == fb)]
    fenlei = v['分类'].iloc[0]
    work = bgs_work(a, b, c, d, e, f, v['平时上班'].iloc[0])
    overtime = bgs_overtime(a, b, c, d, e, f, v['平时加班'].iloc[0])
    late = bgs_late(a, b, c, d, e, f, v['迟到'].iloc[0])
    return fenlei, work, overtime, late


def t(v):
    return datetime.strptime(v, "%H:%M")


def bgs_work(a, b, c, d, e, f, value):
    if value == 0:
        return timedelta(0).seconds
    elif value == '(12:20-8:50)+(d1-13:40)':
        return ((t('12:20') - t('8:50')) + (d[-1] - t('13:40'))).seconds
    elif value == '(12:20-b0)+(d1-13:40)':
        return ((t('12:20') - b[0]) + (d[-1] - t('13:40'))).seconds
    elif value == '(18:00-13:40)+(12:20-8:50)':
        return ((t('18:00') - t('13:40')) + (t('12:20') - t('8:50'))).seconds
    elif value == '12:20-8:50':
        return (t('12:20') - t('8:50')).seconds
    elif value == '12:20-b0':
        return (t('12:20') - b[0]).seconds
    elif value == '18:00-13:40':
        return (t('18:00') - t('13:40')).seconds
    elif value == '18:00-d0':
        return (t('18:00') - d[0]).seconds
    elif value == 'b1-8:50':
        return (b[-1] - t('8:50')).seconds
    elif value == 'b1-b0':
        return (b[-1] - b[0]).seconds
    elif value == 'd1-13:40':
        return (d[-1] - t('13:40')).seconds
    elif value == 'd1-d0':
        return (d[-1] - d[0]).seconds
    elif value == '(12:20-b0)+(18:00-13:40)':
        return ((t('12:20') - b[0]) + (t('18:00') - t('13:40'))).seconds


def bgs_overtime(a, b, c, d, e, f, value):
    if value == 0:
        return 0
    elif value == 'f1-18:30':
        return (f[-1] - t('18:30')).seconds
    elif value == 'f1-f0':
        return (f[-1] - f[0]).seconds


def bgs_late(a, b, c, d, e, f, value):
    # print(value)
    if value == 0:
        return 0
    elif value == '(12:20-8:50)+(d0-13:40)':
        return ((t('12:20') - t('8:50')) + (d[0] - t('13:40'))).seconds
    elif value == '(18:00-13:40)+(12:20-8:50)':
        return ((t('18:00') - t('13:40')) + (t('12:20') - t('8:50'))).seconds
    elif value == '12:20-8:50':
        return (t('12:20') - t('8:50')).seconds
    elif value == 'b0-8:50':
        return (b[0] - t('8:50')).seconds


def cx(s):  # chanxian
    a, b, c, d, e = cat(s.split(' '), '0:00 8:31 12:30 13:31 18:00 23:59'.split(' '), how='[)')
    al, bl, cl, dl, el = [len(i) for i in [a, b, c, d, e]]
    ab, bb, cb, db, eb = [0 if len(i) == 0 else 1 for i in [a, b, c, d, e]]
    v = cx_config[(cx_config['a'] == ab) & (cx_config['b'] == bb) & (cx_config['c'] == cb) & (cx_config['d'] == db) & (
            cx_config['e'] == eb)]
    fenlei = v['分类'].iloc[0]
    work = cx_work(a, b, c, d, e, v['平时上班'].iloc[0])
    overtime = cx_overtime(a, b, c, d, e, v['平时加班'].iloc[0])
    late = cx_late(a, b, c, d, e, v['迟到'].iloc[0])
    return fenlei, work, overtime, late


def cx_work(a, b, c, d, e, value):
    if value == 0:
        return timedelta(0).seconds
    elif value == '(12:30-8:30)+(18:00-13:00)':
        return ((t('12:30') - t('8:30')) + (t('18:00') - t('13:00'))).seconds
    elif value == '(12:30-8:30)+(d1-13:30)':
        return ((t('12:30') - t('8:30')) + (d[-1] - t('13:30'))).seconds
    elif value == '(18:00-13:40)+(12:20-8:50)':
        return ((t('18:00') - t('13:40')) + (t('12:20') - t('8:50'))).seconds
    elif value == '(12:30-b0)+(d1-13:30)':
        return ((t('12:30') - b[0]) + (d[-1] - t('13:30'))).seconds
    elif value == '(12:30-b1)+(18:00-13:30)':
        return ((t('12:30') - b[-1]) + (t('18:00') - t('13:30'))).seconds
    elif value == '12:30-8:30':
        return (t('12:30') - t('8:30')).seconds
    elif value == '12:30-b0':
        return (t('12:30') - b[0]).seconds
    elif value == '18:00-13:30':
        return (t('18:00') - t('13:30')).seconds
    elif value == '18:00-d0':
        return (t('18:00') - d[0]).seconds
    elif value == 'b1-8:30':
        return (b[-1] - t('8:30')).seconds
    elif value == 'b1-b0':
        return (b[-1] - b[0]).seconds
    elif value == 'd1-13:30':
        return (d[-1] - t('13:30')).seconds


def cx_overtime(a, b, c, d, e, value):
    if value == 0:
        return 0
    elif value == 'e1-18:00':
        return (e[-1] - t('18:00')).seconds
    else:
        raise ValueError('加班计算异常')


def cx_late(a, b, c, d, e, value):
    if value == 0:
        return 0
    elif value == '(12:30-8:30)+(d0-13:30)':
        return ((t('12:30') - t('8:30')) + (d[0] - t('13:30'))).seconds
    elif value == '12:30-8:30':
        return (t('12:30') - t('8:30')).seconds
    elif value == 'b0-8:30':
        return (b[0] - t('8:30')).seconds
    else:
        raise ValueError('迟到计算异常')


def main():
    app = QtWidgets.QApplication(sys.argv)  # [固定写法]实例化应用，sys.argv是一个从程序外部获取参数的桥梁
    app.setWindowIcon(QIcon("PO.ico"))  # 设置界面左上角图标
    gui = MainGUI()  # [固定写法]调用MainGUI类
    gui.show()  # [固定写法]显示窗口
    sys.exit(app.exec_())  # [固定写法]调用exec_()方法运行程序。sys.exit()用于程序的正常退出


# [固定写法]执行主函数main()
if __name__ == '__main__':
    main()
