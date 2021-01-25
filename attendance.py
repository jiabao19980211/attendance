import pandas as pd
from datetime import datetime, timedelta
import datetime
from datetime import datetime


urlinput = input("请输入file文件地址:")
urloutput = input("请输入输出文件路径：")
Office_config = pd.read_excel("config.xlsx", 'Sheet3', header=4, engine='openpyxl')
ServiceLine_config = pd.read_excel("config.xlsx", 'Sheet2', header=4, engine='openpyxl')



def list_FllRegisttimeIntoReftimelist(lis_s, cats, how="()"):
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


def Office(s):
    a, b, c, d, e, f = list_FllRegisttimeIntoReftimelist(s.split(' '), '0:00 8:51 12:20 13:41 18:00 18:31 23:59'.split(' '), how='[)')
    al, bl, cl, dl, el, fl = [len(i) for i in [a, b, c, d, e, f]]
    ab, bb, cb, db, eb, fb = [0 if len(i) == 0 else 1 for i in [a, b, c, d, e, f]]
    v = Office_config[
        (Office_config['a'] == ab) & (Office_config['b'] == bb) & (Office_config['c'] == cb) & (Office_config['d'] == db) & (
                Office_config['e'] == eb) & (Office_config['f'] == fb)]
    fenlei = v['分类'].iloc[0]
    work = Office_work(a, b, c, d, e, f, v['平时上班'].iloc[0])
    overtime = Office_overtime(a, b, c, d, e, f, v['平时加班'].iloc[0])
    late = Office_late(a, b, c, d, e, f, v['迟到'].iloc[0])
    return fenlei, work, overtime, late


def t(v):
    return datetime.strptime(v, "%H:%M")


def Office_work(a, b, c, d, e, f, value):
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


def Office_overtime(a, b, c, d, e, f, value):
    if value == 0:
        return 0
    elif value == 'f1-18:30':
        return (f[-1] - t('18:30')).seconds
    elif value == 'f1-f0':
        return (f[-1] - f[0]).seconds


def Office_late(a, b, c, d, e, f, value):
    print(value)
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


def ServiceLine(s):
    a, b, c, d, e = list_FllRegisttimeIntoReftimelist(s.split(' '), '0:00 8:31 12:30 13:31 18:00 23:59'.split(' '), how='[)')
    al, bl, cl, dl, el = [len(i) for i in [a, b, c, d, e]]
    ab, bb, cb, db, eb = [0 if len(i) == 0 else 1 for i in [a, b, c, d, e]]
    v = ServiceLine_config[(ServiceLine_config['a'] == ab) & (ServiceLine_config['b'] == bb) & (ServiceLine_config['c'] == cb) & (ServiceLine_config['d'] == db) & (
            ServiceLine_config['e'] == eb)]
    fenlei = v['分类'].iloc[0]
    work = ServiceLine_work(a, b, c, d, e, v['平时上班'].iloc[0])
    overtime = ServiceLine_overtime(a, b, c, d, e, v['平时加班'].iloc[0])
    late = ServiceLine_late(a, b, c, d, e, v['迟到'].iloc[0])
    return fenlei, work, overtime, late


def ServiceLine_work(a, b, c, d, e, value):
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


def ServiceLine_overtime(a, b, c, d, e, value):
    if value == 0:
        return 0
    elif value == 'e1-18:00':
        return (e[-1] - t('18:00')).seconds
    else:
        raise ValueError('加班计算异常')


def ServiceLine_late(a, b, c, d, e, value):
    print(a, b, c, d, e)
    print(value)
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


def zong():
    df = pd.read_excel(urlinput + "\\file.xlsx", engine='openpyxl')
    df['len'] = df.dropna(subset=['时间'])['时间'].str.split(' ').apply(len)
    df2 = df[df['len'] >= 1]
    Office_df = df2[df2['len'] < 4]
    ServiceLine_df = df2[df2['len'] >= 4]
    Office_df['状态'], Office_df['平时上班'], Office_df['平时加班'], Office_df['迟到'] = zip(*Office_df['时间'].apply(Office))
    ServiceLine_df['状态'], ServiceLine_df['平时上班'], ServiceLine_df['平时加班'], ServiceLine_df['迟到'] = zip(*ServiceLine_df['时间'].apply(ServiceLine))
    Office_df['平时上班'] = Office_df['平时上班'] / 3600
    Office_df['平时加班'] = Office_df['平时加班'] / 3600
    Office_df['迟到'] = Office_df['迟到'] / 60
    ServiceLine_df['平时上班'] = ServiceLine_df['平时上班'] / 3600
    ServiceLine_df['平时加班'] = ServiceLine_df['平时加班'] / 3600
    ServiceLine_df['迟到'] = ServiceLine_df['迟到'] / 60
    Office_df.to_excel(urloutput + "\\ 办公室人员统计.xlsx", 'w+')
    ServiceLine_df.to_excel(urloutput + "\\产线人员统计.xlsx", 'w+')


def main():
    zong()


if __name__ == '__main__':
    main()
