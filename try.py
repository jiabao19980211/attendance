import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
import pandas as pd
from datetime import datetime, timedelta
import datetime
from datetime import datetime

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
        self.input = QtWidgets.QLineEdit()  # 文本编辑框1
        self.input_btn = QtWidgets.QPushButton("选择文件(仅支持xlsx文件)")  # 按钮1及名称
        self.output = QtWidgets.QLineEdit()  # 文本编辑框2
        self.output_btn = QtWidgets.QPushButton("选择输出文件夹")  # 按钮2及名称
        self.show_result = QtWidgets.QListWidget()  # 列表控件
        self.run = QtWidgets.QPushButton("执行转换")  # 按钮3及名称
        # 部件的位置设置
        self.main_widget_layout.addWidget(self.input, 0, 0, 1, 2)  # 文本编辑框1放在第1行1列，占用1行2列
        self.main_widget_layout.addWidget(self.input_btn, 0, 2, 1, 1)  # 按钮1放在第1行第3列，占用1行1列
        self.main_widget_layout.addWidget(self.output, 1, 0, 1, 2)
        self.main_widget_layout.addWidget(self.output_btn, 1, 2, 1, 1)
        self.main_widget_layout.addWidget(self.run, 2, 2, 1, 1)
        self.main_widget_layout.addWidget(self.show_result, 3, 0, 3, 3)

        self.setCentralWidget(self.main_widget)  # 设置QMainWindow的中心窗口

        self.input_btn.clicked.connect(self.Choice_file_input)  # 将"选择输入文件夹"按钮绑定Choice_dir_input函数
        self.output_btn.clicked.connect(self.Choice_dir_output)  # 将"选择输出文件夹"按钮绑定Choice_dir_output函数
        self.run.clicked.connect(self.Summary_data)  # “执行汇总”按钮绑定Summary_data函数
        self.show_result.itemClicked.connect(self.show_select_item)

    def show_select_item(self):
        self.show_result.addItem('功能有待完善')



    def Choice_file_input(self):
        filename, filetype = QFileDialog.getOpenFileName(self, "选取文件", "D:\\", "*.xlsx")
        self.input.setText(filename)

    def Choice_dir_output(self):
        dir_path = QtWidgets.QFileDialog.getExistingDirectory(self, "请选择文件夹路径", "D:\\")
        self.output.setText(dir_path)

    def Summary_data(self):
        urlinput = self.input.text()
        urloutput = self.output.text()
        strInputFileName = os.path.basename(urlinput)[0:-5]  # urlinput.split()#/我想我应该是file吧.xlsx#拿到输入文件名
        df = pd.read_excel(urlinput, engine='openpyxl')
        df['len'] = df.dropna(subset=['时间'])['时间'].str.split(' ').apply(len)
        df2 = df[df['len'] >= 1]
        bgs_df = df2[df2['len'] < 4]
        cx_df = df2[df2['len'] >= 4]
        bgs_df['状态'], bgs_df['平时上班'], bgs_df['平时加班'], bgs_df['迟到'] = zip(*bgs_df['时间'].apply(bgs))
        cx_df['状态'], cx_df['平时上班'], cx_df['平时加班'], cx_df['迟到'] = zip(*cx_df['时间'].apply(cx))
        bgs_df['平时上班'] = bgs_df['平时上班'] / 3600
        bgs_df['平时加班'] = bgs_df['平时加班'] / 3600
        bgs_df['迟到'] = bgs_df['迟到'] / 60
        cx_df['平时上班'] = cx_df['平时上班'] / 3600
        cx_df['平时加班'] = cx_df['平时加班'] / 3600
        cx_df['迟到'] = cx_df['迟到'] / 60

        dt = datetime.now().strftime("%Y%m%d%H%M%S%f")  # 创建一个datetime类
        time = str(dt)

        str1 = urloutput + "/" + strInputFileName + time
        print(str1)
        bgs_df.to_excel(str1 + "办公室人员统计.xlsx", 'w+')
        cx_df.to_excel(str1 + "产线人员统计.xlsx", 'w+')

        self.show_result.addItem("------------------------------------------")
        self.show_result.addItem("文件生成成功！！")
        self.show_result.addItem("输入地址为：" + urlinput)
        self.show_result.addItem("输出地址为：" + urloutput)
        self.show_result.addItem("文件1为：" + strInputFileName + time + "办公室人员统计.xlsx")
        self.show_result.addItem("文件2为：" + strInputFileName + time + "产线人员统计.xlsx")


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


def cx(s):
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
    # print(a, b, c, d, e)
    # print(value)
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
