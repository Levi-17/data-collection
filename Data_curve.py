#!/usr/bin/python
# -*- coding: UTF-8 -*-


'''
Case Info:
    Name:This is a Case Title
    Pre:
        1.Get data
        2.Make a graph
    TestStep:
        1.Fetching the data
        2.Import data into excel
    Author:levi.zhang
    Topo:Win
'''


import xlsxwriter
import re
import os

# 新建EXCEL文件
nwb=xlsxwriter.Workbook('Smart Attributes-ZTE6.xlsx')

def make_sheet(nwb,SN_num):
    # 新建表格
    nws=nwb.add_worksheet(SN_num)

    # 定义首行
    row1=['COUNT','ID#1-Raw_Read_Error_Rate','ID#5-Reallocated_or_Ct','ID#9-Power_On_Hours',
          'ID#12-Power_Cycle_Count','ID#160-Uncorrectable Sector Count','ID#161-Valid Spare Blocks',
          'ID#163-Initial Invalid Blocks','ID#164-Total Erase Count','ID#165-Maximum Erase Count',
          'ID#166-Minimum Erase Count','ID#167-Average Erase Count','ID#169-Percentage Lifetime Remaining',
          'ID#194-Temperature_Celsius','ID#195-Hardware ECC recovered','ID#196-Reallocation Event Count',
          'ID#198-Uncorrectable Sector Count Offline','ID#199-UDMA CRC Error Count','ID#241-Host Write Sector Count',
          'ID#242-Read Sector Count','ID#245-Flash Write Count']
    row2='COUNT'
    i=0

    for j in range(0,len(row1)):
        nws.write(i,j,row1[j])

    nws.write(i+1,0,row2)

    return nws


def get_data(file,nws):
    count=0
    i=1
    del_list=[]
    # 需获取数据的id列表
    id_list=['1','5','9','12','160','161','163','164','165','166','167','169','194','195','196','198','199','241','242','245']
    with open(file,'r') as f:
        for line in f.readlines():
            if line=='\n':
                continue
            line=line.lstrip()
            line=line.split()
            del line[2:9]
            for id in range(0,len(id_list)):
                if line[0] == id_list[id]:
                    if id_list[id] == '245':
                        nws.write(i, 20, int(line[2]))
                        i+=1
                    elif id_list[id] == '194':
                        for num in line:
                            if re.search(r'\d+/\d+\)',num):
                                del_list.append(num)
                        line.remove("(Min/Max")
                        for del_num in del_list:
                            del_list.remove(del_num)
                        nws.write(i, 13, int(line[2]))
                    elif id_list[id]=='1':
                        count += 1
                        nws.write(i, 1, int(line[2]))


                    nws.write(i, id+1, int(line[2]))
                    nws.write(i,0,count)

    return count



            # if line[0]==id_list[0]:
            #     count+=1
            #     nws.write(i,1,line[2])
            #
            # if line[0] == id_list[1]:
            #     nws.write(i, 2, line[2])
            #
            # if line[0] == id_list[2]:
            #     nws.write(i, 3, line[2])
            #
            # if line[0] == id_list[3]:
            #     nws.write(i, 4, line[2])
            #

            # print(type(line[0]))
            # for id in id_list:
            #     if line[0]==id:
            #         if id=='1':
            #             count += 1
                    # if id == '161':
                    #     line[1] = 'Valid Spare Blocks'
                    # if id == '167':
                    #     line[1] = 'Average Erase Count'
                    # if id == '169':
                    #     line[1] = 'Percentage Lifetime Remaining'
                    # if id == '194':
                    #     line.remove("(Min/Max")
                    #     line.remove("10/50)")
                    # if id == '241':
                    #     line[1] = 'Host Write Sector Count'
                    # if id == '245':
                    #     line[1] = 'Flash Write Count'

                    # for j in range(1,len(id_list)+1):
                    #     nws.write(i, j, line[2])
                    #     nws.write(0,i,count)


def add_chart1(nws,count):
    # 创建第一个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第一个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$B$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$B$2:$B$'+str(count+1),
        'line': {'color': 'green'},
    })


    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$B$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X10', chart_col, {'x_offset': 50, 'y_offset': 40,})


def add_chart2(nws,count):
    # 创建第二个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第二个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$C$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$C$2:$C$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$C$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X40', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart3(nws,count):
    # 创建第三个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第第三个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$D$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$D$2:$D$'+str(count+1),
        'line': {'color': 'green'},
    })


    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$D$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X70', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart4(nws,count):
    # 创建第四个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第四个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$E$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$E$2:$E$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$E$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X100', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart5(nws,count):
    # 创建第五个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第五个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$F$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$F$2:$F$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$F$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X130', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart6(nws,count):
    # 创建第六个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第六个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$G$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$G$2:$G$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$G$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X160', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart7(nws,count):
    # 创建第七个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第七个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$H$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$H$2:$H$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$H$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X190', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart8(nws,count):
    # 创建第八个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第八个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$I$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$I$2:$I$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$I$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X220', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart9(nws,count):
    # 创建第九个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第九个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$J$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$J$2:$J$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$J$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X250', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart10(nws,count):
    # 创建第十个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$K$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$K$2:$K$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$K$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X280', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart11(nws,count):
    # 创建第十一个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十一个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$L$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$L$2:$L$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$L$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X310', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart12(nws,count):
    # 创建第十二个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十二个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$M$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$M$2:$M$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$M$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X340', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart13(nws,count):
    # 创建第十三个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十二个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$N$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$N$2:$N$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$N$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X370', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart14(nws,count):
    # 创建第十四个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十四个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$O$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$O$2:$O$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$O$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X400', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart15(nws,count):
    # 创建第十五个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十五个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$P$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$P$2:$P$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$P$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X430', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart16(nws,count):
    # 创建第十六个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十六个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$Q$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$Q$2:$Q$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$Q$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X460', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart17(nws,count):
    # 创建第十七个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十七个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$R$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$R$2:$R$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$R$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X490', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart18(nws,count):
    # 创建第十八个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十八个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$S$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$S$2:$S$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$S$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X520', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart19(nws,count):
    # 创建第十九个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第十九个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$T$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$T$2:$T$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$T$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X550', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart20(nws,count):
    # 创建第二十个折线图(line chart)
    chart_col = nwb.add_chart({'type': 'line'})

    # 配置第二十个系列数据
    chart_col.add_series({
        'name': '='+str(nws.name)+'!$U$1',
        'categories': '='+str(nws.name)+'!$A$2:$A$'+str(count+1),
        'values': '='+str(nws.name)+'!$U$2:$U$'+str(count+1),
        'line': {'color': 'green'},
    })

    # 设置图表的title 和 x，y轴信息
    chart_col.set_title({'name': '='+str(nws.name)+'!$U$1'})
    chart_col.set_x_axis({'name': 'Print Count'})
    chart_col.set_y_axis({'name': 'Data Value'})

    # 设置图表的风格
    chart_col.set_style(1)

    # 把图表插入到worksheet并设置偏移
    nws.insert_chart('X580', chart_col, {'x_offset': 50, 'y_offset': 40})


def add_chart(nws,count):
    add_chart1(nws,count)
    add_chart2(nws,count)
    add_chart3(nws,count)
    add_chart4(nws,count)
    add_chart5(nws,count)
    add_chart6(nws,count)
    add_chart7(nws,count)
    add_chart8(nws,count)
    add_chart9(nws,count)
    add_chart10(nws,count)
    add_chart11(nws,count)
    add_chart12(nws,count)
    add_chart13(nws,count)
    add_chart14(nws,count)
    add_chart15(nws,count)
    add_chart16(nws,count)
    add_chart17(nws,count)
    add_chart18(nws,count)
    add_chart19(nws,count)
    add_chart20(nws,count)



def get_dir_list(path):
    # 获取要读取数据的文件，并将文件目录下的子文件输出到一个列表
    dir_list=[]
    dir_p=os.listdir(path)
    for dir in dir_p:
        dir_path = os.path.join(path, dir)
        dir_list.append(dir_path)
    return dir_list


def get_SN_list(path):

    # 首先遍历当前目录所有文件及文件夹
    file_list = os.listdir(path)
    # 准备循环判断每个元素是否是文件夹还是文件，是文件的话，把名称传入list，是文件夹的话，递归
    # for file in file_list:
        # 利用os.path.join()方法取得路径全名，并存入file_path变量
        # file_path = os.path.join(path, file)
        # if os.path.isfile(file_path):
        #     if re .search(r'sd.+',file_path):
        #         File_list.append(file_path)
        #     elif re.search(r'log.*', file_path):
        #         with open(file_path,'r')as f:
        #             for line in f.readlines():
        #                 if re.search(r"\'sn\':(.+?),",line):
        #                     data=re.findall(r"\'sn\':(.+?),",line)
        #                     SN_list.append(str(data))
    sn_dict=dict()
    sd_dict=dict()
    for file in file_list:
        file_path=os.path.join(path,file)
        if os.path.isfile(file_path):
            sds=re.findall("(sd\w+)",file_path)
            if sds:
                sd_dict[sds[0]]=file_path
            else:
                # with open(file_path, 'r')as f:
                #     for line in f.readlines():
                #         sns=re.findall(r"\'sn\':(.+?),",line)
                with open(file_path,"r") as f:
                    for line in f :
                        sns=re.findall("Test disk: {[\s,\S]+?/dev/(sd\w+)[\s,\S]+?sn\':\s*\'([\S,\s]+?)\'",line)
                        for ks in sns:
                            if ks[1] not in sn_dict:
                                sn_dict[ks[1]]=ks[0]
    for k,v in sn_dict.items():
        sn_dict[k]=sd_dict[v]

    return sn_dict



def main():
    global nwb
    dir_list = get_dir_list('C:/Users/bwhq-rd-1783/Desktop/ZTE/6')
    for dir in dir_list:
        # dir_name = os.path.basename(dir)
        # dir_name = dir_name.split('.')[0]
        dict=get_SN_list(dir)
        for k,v in dict.items():
            nws=make_sheet(nwb,k)
            c=get_data(v,nws)
            add_chart(nws,c)


    nwb.close()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)





