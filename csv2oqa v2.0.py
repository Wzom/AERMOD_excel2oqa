#coding:utf-8 

##---------------------------------##
###作者：牛钰功
###联系方式：niuyugong@outlook.com
###日期：2020-08-14
##---------------------------------##

import numpy as np
import pandas as pd
import sys 


#输入参数定义
input_file_name = sys.argv[1]
input_sheet_name = sys.argv[2]
output_file_name = sys.argv[3]

#读入execl文件
data = pd.read_excel(input_file_name, input_sheet_name)


#检查缺失项
#cols_refer_list：全表头列表
#data_cols_list：execl输入列表头列表
cols_refer_list = ['年','月','日','小时','降水量','海平面气压','测点气压','云层高度','总云量','低云量','1层云层状况','2层云层状况','3层云层状况','4层云层状况','5层云层状况','6层云层状况','天气代码（近地面）','天气代码','ASOS天气','ASOS高度','水平能见度','干球温度','湿球温度','露点温度','相对湿度','风向','风速']
data_cols_list = list(data)
#cols_list：缺失列表头list
cols_list = cols_refer_list + data_cols_list
for x in cols_refer_list:
    if cols_list.count(x) == 2:
        cols_list.remove(x)
        cols_list.remove(x)


#定义缺失列参考值
refer_series = pd.Series(['-9', '99999', '99999', '999', '9999', '9999', '09999', '09999', '09999', '09999', '09999', '09999', '9999', '9999', '99', '999', '99999', '999', '999', '999', '999', '99', '-9999'], index=['降水量','海平面气压','测点气压','云层高度','总云量','低云量','1层云层状况','2层云层状况','3层云层状况','4层云层状况','5层云层状况','6层云层状况','天气代码（近地面）','天气代码','ASOS天气','ASOS高度','水平能见度','干球温度','湿球温度','露点温度','相对湿度','风向','风速'])

#refer：缺失列参考值列表
refer = []
for x in cols_list:
    refer.append(refer_series[x])

#refer_val_series：缺失列参考值Series
refer_val_series = pd.Series(refer, index=cols_list)


#execl输入列数值预处理
#data_array：execl输入值array
data_array = data.values
[rows,cols] = data_array.shape

#处理"/"
for x in range(rows):
    for y in range(cols):
        if data_array[x][y] == '/':
            data_array[x][y] = np.NAN

#定义时间序列数组
yy = []
mm = []
dd = []
hh = []

#超出范围值处理、数值处理、风向转换及时间序列缺失填补
for x in data_cols_list:
    for i in range(rows):
        if x == '年':
            cols_yy = data_cols_list.index(x)
            if data_array[i][cols_yy]<1900 or data_array[i][cols_yy]>3000 or data_array[i][cols_yy] is np.NAN:
                if i != 0:
                    data_array[i][cols_yy] = data_array[i-1][cols_yy]
                else:
                    data_array[i][cols_yy] = data_array[i+1][cols_yy]
            yy.append(data_array[i][cols_yy])
        elif x == '月':
            cols_mm = data_cols_list.index(x)
            if data_array[i][cols_mm]<1 or data_array[i][cols_mm]>12 or data_array[i][cols_mm] is np.NAN:
                if i != 0:
                    data_array[i][cols_mm] = data_array[i-1][cols_mm]
                else:
                    data_array[i][cols_mm] = data_array[i+1][cols_mm]
            mm.append(data_array[i][cols_mm])
        elif x == '日':
            cols_dd = data_cols_list.index(x)
            if data_array[i][cols_dd]<1 or data_array[i][cols_dd]>31 or data_array[i][cols_dd] is np.NAN:
                if i != 0:
                    data_array[i][cols_dd] = data_array[i-1][cols_dd]
                else:
                    data_array[i][cols_dd] = data_array[i+1][cols_dd]
            dd.append(data_array[i][cols_dd])
        elif x == '小时':
            cols_hh = data_cols_list.index(x)
            if data_array[i][cols_dd]<1 or data_array[i][cols_dd]>24 or data_array[i][cols_hh] is np.NAN:
                if i != 0:
                    data_array[i][cols_hh] = data_array[i-1][cols_hh]
                else:
                    data_array[i][cols_hh] = data_array[i+1][cols_hh]
            hh.append(data_array[i][cols_hh])
        elif x == '降水量':
            cols_PRCP = data_cols_list.index(x)
            data_array[i][cols_PRCP] *= 1000
            if data_array[i][cols_PRCP]>25400 or data_array[i][cols_PRCP]<0:
                data_array[i][cols_PRCP] = np.NAN
        elif x == '海平面气压':
            cols_SLVP = data_cols_list.index(x)
            data_array[i][cols_SLVP] *= 10
            if data_array[i][cols_SLVP]>10999 or data_array[i][cols_SLVP]<9000:
                data_array[i][cols_SLVP] = np.NAN
        elif x == '测点气压':
            cols_PRES = data_cols_list.index(x)
            data_array[i][cols_PRES] *= 10
            if data_array[i][cols_PRES]>10999 or data_array[i][cols_PRES]<9000:
                data_array[i][cols_PRES] = np.NAN
        elif x == '云层高度':
            cols_CLHT = data_cols_list.index(x)
            data[i][cols_CLHT] *= 10
            if data[i][cols_CLHT]>300 or data[i][cols_CLHT]<0:
                data[i][cols_CLHT] = np.NAN
        elif x == '总云量':
            cols_TSKCT = data_cols_list.index(x)
            if data_array[i][cols_TSKCT]>1010 or data_array[i][cols_TSKCT]<0:
                data_array[i][cols_TSKCT] = np.NAN
        elif x == '低云量':
            cols_TSKCL = data_cols_list.index(x)
            if data_array[i][cols_TSKCL]>1010 or data_array[i][cols_TSKCL]<0:
                data_array[i][cols_TSKCL] = np.NAN
        elif x == '1层云层状况':
            cols_ALC1 = data_cols_list.index(x)
            if data_array[i][cols_ALC1]>300 or data_array[i][cols_ALC1]<0:
                data_array[i][cols_ALC1] = np.NAN
        elif x == '2层云层状况':
            cols_ALC2 = data_cols_list.index(x)
            if data_array[i][cols_ALC2]>300 or data_array[i][cols_ALC2]<0:
                data_array[i][cols_ALC2] = np.NAN
        elif x == '3层云层状况':
            cols_ALC3 = data_cols_list.index(x)
            if data_array[i][cols_ALC3]>300 or data_array[i][cols_ALC3]<0:
                data_array[i][cols_ALC3] = np.NAN
        elif x == '4层云层状况':
            cols_ALC4 = data_cols_list.index(x)
            if data_array[i][cols_ALC4]>850 or data_array[i][cols_ALC4]<0:
                data_array[i][cols_ALC4] = np.NAN
        elif x == '5层云层状况':
            cols_ALC5 = data_cols_list.index(x)
            if data_array[i][cols_ALC5]>850 or data_array[i][cols_ALC5]<0:
                data_array[i][cols_ALC5] = np.NAN
        elif x == '6层云层状况':
            cols_ALC6 = data_cols_list.index(x)
            if data_array[i][cols_ALC6]>850 or data_array[i][cols_ALC6]<0:
                data_array[i][cols_ALC6] = np.NAN
        elif x == '天气代码（近地面）':
            cols_PWVC = data_cols_list.index(x)
            if data_array[i][cols_PWVC]>98300 or data_array[i][cols_PWVC]<9292:
                data_array[i][cols_PWVC] = np.NAN
        elif x == '天气代码':
            cols_PWTH = data_cols_list.index(x)
            if data_array[i][cols_PWTH]>98300 or data_array[i][cols_PWTH]<9292:
                data_array[i][cols_PWTH] = np.NAN
        elif x == 'ASOS天气':
            cols_ASKY = data_cols_list.index(x)
            if data_array[i][cols_ASKY]>10 or data_array[i][cols_ASKY]<0:
                data_array[i][cols_ASKY] = np.NAN
        elif x == 'ASOS高度':
            cols_ACHT = data_cols_list.index(x)
            data_array[i][cols_ACHT] *= 10
            if data_array[i][cols_ACHT]>888 or data_array[i][cols_ACHT]<0:
                data_array[i][cols_ACHT] = np.NAN
        elif x == '水平能见度':
            cols_HZVS = data_cols_list.index(x)
            data_array[i][cols_HZVS] *= 10
            if data_array[i][cols_HZVS]>1640 or data_array[i][cols_HZVS]<0:
                data_array[i][cols_HZVS] = np.NAN
        elif x == '干球温度':
            cols_TMPD = data_cols_list.index(x)
            data_array[i][cols_TMPD] *= 10
            if data_array[i][cols_TMPD]>350 or data_array[i][cols_TMPD]<-300:
                data_array[i][cols_TMPD] = np.NAN
        elif x == '湿球温度':
            cols_TMPV = data_cols_list.index(x)
            data_array[i][cols_TMPV] *= 10
            if data_array[i][cols_TMPV]>350 or data_array[i][cols_TMPV]<-650:
                data_array[i][cols_TMPV] = np.NAN
        elif x == '露点温度':
            cols_DPTP = data_cols_list.index(x)
            data_array[i][cols_DPTP] *= 10
            if data_array[i][cols_DPTP]>350 or data_array[i][cols_DPTP]<-650:
                data_array[i][cols_DPTP] = np.NAN
        elif x == '相对湿度':
            cols_RHUM = data_cols_list.index(x)
            if data_array[i][cols_RHUM]>100 or data_array[i][cols_RHUM]<0:
                data_array[i][cols_RHUM] = np.NAN
        elif x == '风向':
            cols_WDIR = data_cols_list.index(x)
            if data_array[i][cols_WDIR]=='N':
                data_array[i][cols_WDIR] = 0
            elif data_array[i][cols_WDIR]=='NNE':
                data_array[i][cols_WDIR] = 2
            elif data_array[i][cols_WDIR]=='NE':
                data_array[i][cols_WDIR] = 6
            elif data_array[i][cols_WDIR]=='ENE':
                data_array[i][cols_WDIR] = 7
            elif data_array[i][cols_WDIR]=='E':
                data_array[i][cols_WDIR] = 9
            elif data_array[i][cols_WDIR]=='ESE':
                 data_array[i][cols_WDIR] = 11
            elif data_array[i][cols_WDIR]=='SE':
                data_array[i][cols_WDIR] = 14
            elif data_array[i][cols_WDIR]=='SSE':
                 data_array[i][cols_WDIR] = 16
            elif data_array[i][cols_WDIR]=='S':
                data_array[i][cols_WDIR] = 18
            elif data_array[i][cols_WDIR]=='SSW':
                data_array[i][cols_WDIR] = 20
            elif data_array[i][cols_WDIR]=='SW':
                data_array[i][cols_WDIR] = 23
            elif data_array[i][cols_WDIR]=='WSW':
                data_array[i][cols_WDIR] = 25
            elif data_array[i][cols_WDIR]=='W':
                    data_array[i][cols_WDIR] = 27
            elif data_array[i][cols_WDIR]=='WNW':
                data_array[i][cols_WDIR] = 29
            elif data_array[i][cols_WDIR]=='NW':
                data_array[i][cols_WDIR] = 32
            elif data_array[i][cols_WDIR]=='NNW':
                data_array[i][cols_WDIR] = 34
            elif data_array[i][cols_WDIR]=='C':
                data_array[i][cols_WDIR] = 0
        elif x == '风速':
            cols_WSPD = data_cols_list.index(x)
            data_array[i][cols_WSPD] *= 10
            if data_array[i][cols_WSPD]>500 or data_array[i][cols_WSPD]<0:
                data_array[i][cols_WSPD] = np.NAN

#时间序列转换
yymmddhh_array = np.empty((rows),dtype=int)
for i in range(rows):
    if yy[i] == yy[i]:
        if mm[i] == mm[i]:
            if dd[i] == dd[i]:
                if hh[i] == hh[i]:
                    yymmddhh_array[i] = yy[i]%100 * 1000000 + mm[i] * 10000 + dd[i] * 100 + hh[i]
                else:
                    yymmddhh_array[i] = 0
            else:
                yymmddhh_array[i] = 0
        else:
            yymmddhh_array[i] = 0
    else:
        yymmddhh_array[i] = 0
yymmddhh = yymmddhh_array.tolist()

#删除输入数据中的年月日时
for x in data_cols_list:
    if x == '年':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
        data_cols_list.remove(x)
    elif x == '月':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
        data_cols_list.remove(x)
    elif x == '日':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
        data_cols_list.remove(x)
    elif x == '小时':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
        data_cols_list.remove(x)

#时间序列转换
yymmddhh_array = np.empty((rows),dtype=int)
for i in range(rows):
    if yy[i] == yy[i]:
        if mm[i] == mm[i]:
            if dd[i] == dd[i]:
                if hh[i] == hh[i]:
                    yymmddhh_array[i] = yy[i]%100 * 1000000 + mm[i] * 10000 + dd[i] * 100 + hh[i]
                else:
                    yymmddhh_array[i] = 0
            else:
                yymmddhh_array[i] = 0
        else:
            yymmddhh_array[i] = 0
    else:
        yymmddhh_array[i] = 0
yymmddhh = yymmddhh_array.tolist()

#删除输入数据中的年月日时
for x in data_cols_list:
    if x == '年':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
    elif x == '月':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
    elif x == '日':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
    elif x == '小时':
        data_array = np.delete(data_array, data_cols_list.index(x), axis=1)
for x in data_cols_list:
    if x == '小时':
        data_cols_list.remove(x)
for x in data_cols_list:
    if x == '年':
        data_cols_list.remove(x)
for x in data_cols_list:
    if x == '月':
        data_cols_list.remove(x)
for x in data_cols_list:
    if x == '日':
        data_cols_list.remove(x)
    

#缺失值处理
[r, c] = data_array.shape
for i in range(c):
    if data_array[0][i] is np.NAN:
        data_array[0][i] = data_array[1][i]
data_df = pd.DataFrame(data_array, columns=data_cols_list, dtype='float')
#线性插值
#nearest：最邻近插值法
#zero：阶梯插值
#slinear、linear：线性插值
#quadratic、cubic：2、3阶B样条曲线插值
data_nonan_df = data_df.interpolate(method='linear', limit_direction='forward', axis=0).astype('int')


#输出OQA文件
try:
    file = open(output_file_name, 'w+')
    file.write("*% SURFACE\n")
    file.write("*      XDATES    %d/%d/%d  TO  %d/%d/%d\n" % (yy[0], mm[0], dd[0], yy[-1], mm[-1], dd[-1]))
    file.write("*@     LOCATION  00000001  27.7N  102.15E  0  34\n")
    file.write("*** EOH: END OF SURFACE QA HEADERS\n")
    for i in range(r):
        file.write(str(yymmddhh[i]))
        file.write("  ")
        if data_cols_list.count('降水量'):
            file.write(str(data_nonan_df['降水量'][i]))
        else:
            file.write("-9")
        file.write("  ")
        if data_cols_list.count('海平面气压'):
            file.write(str(data_nonan_df['海平面气压'][i]))
        else:
            file.write("99999")
        file.write("  ")
        if data_cols_list.count('测点气压'):
            file.write(str(data_nonan_df['测点气压'][i]))
        else:
            file.write("99999")
        file.write("  ")
        if data_cols_list.count('云层高度'):
            file.write(str(data_nonan_df['云层高度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('总云'):
            file.write(str(data_nonan_df['总云'][i]))
        else:
            file.write("9999")
        file.write("  ")
        if data_cols_list.count('1层云层状况'):
            file.write(str(data_nonan_df['1层云层状况'][i]))
        else:
            file.write("09999")
        file.write("  ")
        if data_cols_list.count('2层云层状况'):
            file.write(str(data_nonan_df['2层云层状况'][i]))
        else:
            file.write("09999")
        file.write("  ")
        if data_cols_list.count('3层云层状况'):
            file.write(str(data_nonan_df['3层云层状况'][i]))
        else:
            file.write("09999")
        file.write("  ")
        if data_cols_list.count('4层云层状况'):
            file.write(str(data_nonan_df['4层云层状况'][i]))
        else:
            file.write("09999")
        file.write("  ")
        if data_cols_list.count('5层云层状况'):
            file.write(str(data_nonan_df['5层云层状况'][i]))
        else:
            file.write("09999")
        file.write("\n")
        if data_cols_list.count('6层云层状况'):
            file.write(str(data_nonan_df['6层云层状况'][i]))
        else:
            file.write("09999")
        file.write("  ")
        if data_cols_list.count('天气代码（近地面）'):
            file.write(str(data_nonan_df['天气代码（近地面）'][i]))
        else:
            file.write("9999")
        file.write("  ")
        if data_cols_list.count('天气代码'):
            file.write(str(data_nonan_df['天气代码'][i]))
        else:
            file.write("9999")
        file.write("  ")
        if data_cols_list.count('ASOS天气'):
            file.write(str(data_nonan_df['ASOS天气'][i]))
        else:
            file.write("99")
        file.write("  ")
        if data_cols_list.count('ASOS高度'):
            file.write(str(data_nonan_df['ASOS高度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('水平能见度'):
            file.write(str(data_nonan_df['水平能见度'][i]))
        else:
            file.write("99999")
        file.write("  ")
        if data_cols_list.count('干球温度'):
            file.write(str(data_nonan_df['干球温度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('湿球温度'):
            file.write(str(data_nonan_df['湿球温度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('露点温度'):
            file.write(str(data_nonan_df['露点温度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('相对湿度'):
            file.write(str(data_nonan_df['相对湿度'][i]))
        else:
            file.write("999")
        file.write("  ")
        if data_cols_list.count('风向'):
            file.write(str(data_nonan_df['风向'][i]))
        else:
            file.write("99")
        file.write("  ")
        if data_cols_list.count('风速'):
            file.write(str(data_nonan_df['风速'][i]))
        else:
            file.write("-9999")
        file.write("  ")
        file.write("\'N\'")
        file.write("\n")
finally:
    if file:
        file.close()