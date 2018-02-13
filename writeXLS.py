#coding: utf8
import xlsxwriter

# 创建一个excel文件
workbook = xlsxwriter.Workbook('chart.xls')
# 创建一个工作表对象,sheet栏
worksheet = workbook.add_worksheet('game') #u'数据报表'
# 创建一个图表对象 type:colum(柱状图)
chart = workbook.add_chart({'type':'column'})
# 定义数据表头列表
title = [u'游戏名称',u'星期一',u'星期二',u'星期三',u'星期四',u'星期五',u'星期六',u'星期日',u'平均时长']
# 定义业务名称列表
buname = [u'英雄联盟',u'王者荣耀',u'洛奇英雄转',u'剑灵',u'龙之谷']
# 定义数据
data = [
    [150,152,158,149,155,145,148],
    [89,88,95,93,98,100,99],
    [201,200,198,175,170,198,195],
    [75,77,78,78,74,70,79],
    [88,85,87,90,93,88,84]
]

# 定义format格式对象
format = workbook.add_format()
# 定义format对象单元格边框加粗(1像素)的格式
format.set_border(1)

# 定义format_title格式对象
format_title = workbook.add_format()
format_title.set_border(1)
# 定义format_title对象单元格背景颜色
format_title.set_bg_color('#cccccc')

# 定义format_ave单元格式
format_ave = workbook.add_format()
format_ave.set_border(1)
# 定义format_ave对象单元格数字显示格式(小数点后2位)
format_ave.set_num_format('0.00')

# 将数据,信息写入xls文件
#以行的方式写
worksheet.write_row('A1',title,format_title)
worksheet.write_column('A2',buname,format)
worksheet.write_row('B2',data[0],format)
worksheet.write_row('B3',data[1],format)
worksheet.write_row('B4',data[2],format)
worksheet.write_row('B5',data[3],format)
worksheet.write_row('B6',data[4],format)

# 定义图表数据系列函数
def chart_series(cur_row):
    '''
    绘制柱状图
    :param cur_row:行号String类型
    :return:
    '''
    # 计算(AVERAGE函数)频道周平均流量
    worksheet.write_formula('I'+cur_row,'=AVERAGE(B'+cur_row+':H'+cur_row+')',format_ave)

    # 画图
    chart.add_series({
        'categories': '=game!$B$1:$H$1',  # 将"星期一至星期日"作为图表数据标签(X轴)
        'values': '=game!$B$'+cur_row+':$H$'+cur_row,  # 频道一周所有数据作为数据区域
        'line': {'color': 'black'},  # 线条颜色定义为black
        'name': '=game!$A$'+cur_row  #引用业务名称为图例项
    })


# 数据以2-6行进行图表数据系列函数
for row in range(2,7):
    chart_series(str(row))

#chart.set_table() # 设置x轴表格格式
# chart.set_style(30)  # 设置图表样式

# 设置图表大小
chart.set_size({'width':577,'height':287})
# 设置标题
chart.set_title({'name':u'游戏时长周报图表'})
# 设置y轴(左侧)小标题
chart.set_y_axis({'name': u'h(小时)'})
# chart.set_y_2axis()

#将图表插入在A8单元格
worksheet.insert_chart('A8',chart)

# 关闭xls
workbook.close()