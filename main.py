import time
import threading
import os
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
import pandas as pd

begin_value = 'Sample'  # log数据开头第一个单词，一般为Sample
file_name = '7250-001-cycle-test'  # log文件名
excel_name = 'test'  # excel文件名

'''
根据不同的芯片输入ic_type的值
1: BQ28Z610, 
2: SN27541
若没有相应的芯片类型，则可以把ic_type置为0，然后自定义type0的数据
'''
ic_type = 2

'''
根据log数据中输出值的命名来修改type0中的值
如log上的时间名为ElapsedTime，则把type0中的TimeName改为ElapsedTime
'''
type0 = ['TimeName', 'VoltageName', 'CurrentName', 'RSOCName', 'RCName', 'FCCName']

g_time_flag = 0


def time_count():
    global g_time_flag
    while True:
        if g_time_flag == 0:
            time.sleep(0.1)
            continue
        elif g_time_flag == 1:
            start_time = time.time()
            while g_time_flag:
                print_time = time.time() - start_time
                print("\r%.2fs" % print_time, end='')
                time.sleep(0.01)


class BuildExcel:
    def __init__(self, log_name):
        self.file_path = 'E:/Python/project/log_2_excel/' + file_name + '.log'
        self.excel_path = 'E:/Python/project/log_2_excel/result/' + excel_name + '.xlsx'
        self.log_name = log_name

    def log_to_excel(self):
        file = open(self.file_path, 'r')
        line = file.readlines()
        file.close()

        try:
            os.mkdir('E:/Python/project/log_2_excel/result/')
        except:
            pass

        for i in range(len(line)):
            if begin_value in line[i]:
                begin_num = i
                break
            else:
                continue

        if ',' in line[begin_num]:
            delimiter = ','
        elif '\t' in line[begin_num]:
            delimiter = '\t'

        i = begin_num
        while i < len(line):
            line[i] = line[i].split(delimiter)
            i += 1

        new_line = line[begin_num:]

        for i in range(len(new_line)):
            if i == 0:
                time_num = new_line[i].index(self.log_name[0])
                voltage_num = new_line[i].index(self.log_name[1])
                current_num = new_line[i].index(self.log_name[2])
                rsoc_num = new_line[i].index(self.log_name[3])
                rc_num = new_line[i].index(self.log_name[4])
                fcc_num = new_line[i].index(self.log_name[5])
                new_line[i].extend([' ', 'Time', 'Voltage', 'Current', 'RSOC', 'RC', 'FCC'])

            else:
                ''' 将通讯错误引起的空白值改为0 '''
                if not new_line[i][time_num]:
                    new_line[i][time_num] = 0
                if not new_line[i][voltage_num]:
                    new_line[i][voltage_num] = 0
                if not new_line[i][current_num]:
                    new_line[i][current_num] = 0
                if not new_line[i][rsoc_num]:
                    new_line[i][rsoc_num] = 0
                if not new_line[i][rc_num]:
                    new_line[i][rc_num] = 0
                if not new_line[i][fcc_num]:
                    new_line[i][fcc_num] = 0

                new_line[i].extend([' ',
                                    round(float(new_line[i][time_num]) / 3600, 6),
                                    int(new_line[i][voltage_num]),
                                    abs(int(new_line[i][current_num])),
                                    int(new_line[i][rsoc_num]),
                                    int(new_line[i][rc_num]),
                                    int(new_line[i][fcc_num])
                                    ])
        df = pd.DataFrame(new_line)
        df.to_excel(self.excel_path, header=None, index=False)

    def print_chart(self):
        file = openpyxl.load_workbook(self.excel_path)
        sheet = file.active

        chart_sheet = file.create_chartsheet('Chart1')

        chart = ScatterChart()
        chart.title = 'test'

        xvalue = Reference(sheet, min_row=2, min_col=sheet.max_column - 5,
                           max_row=sheet.max_row, max_col=sheet.max_column - 5)

        for i in range(sheet.max_column-4, sheet.max_column+1):
            yvalue = Reference(sheet, min_row=1, min_col=i,
                               max_row=sheet.max_row, max_col=i)

            series = Series(yvalue, xvalue, title_from_data=True)
            chart.append(value=series)
        chart_sheet.add_chart(chart)
        file.save(self.excel_path)
        file.close()


if __name__ == '__main__':
    if ic_type == 1:
        module_name = ['ElapsedTime', 'Voltage', 'Current', 'RSOC', 'RemCap', 'FullChgCap']
    elif ic_type == 2:
        module_name = ['~Elapsed(s)', 'Voltage', 'AvgCurrent', 'StateofChg', 'RemCap', 'FullChgCap']
    else:
        module_name = type0

    bulid_excel = BuildExcel(module_name)
    time_count_thread = threading.Thread(target=time_count)
    time_count_thread.daemon = True
    time_count_thread.start()

    print('正在将log数据写入excel，请耐心等待...')
    g_time_flag = 1
    bulid_excel.log_to_excel()
    g_time_flag = 0
    print('\n写入完成')

    time.sleep(0.1)

    print('正在绘制图表，请耐心等待...')
    g_time_flag = 1
    bulid_excel.print_chart()
    g_time_flag = 0
    print('\n画图完成')
