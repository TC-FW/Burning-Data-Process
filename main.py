import glob
import time
import threading
import os
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
import pandas as pd

begin_value = 'Sample'  # log数据开头第一个单词，一般为Sample

'''
目前支持的芯片列表如下：
BQ28Z610, BQ40Z50R2, SN27541
若没有相应的芯片类型，则可以把custom_type置为True，然后自定义custom_name的数据
'''
custom_type = False

'''
根据log数据中输出值的命名来修改type0中的值
如log上的时间名为ElapsedTime，则把custom_name中的TimeName改为ElapsedTime
'''
custom_name = ['TimeName', 'VoltageName', 'CurrentName', 'RSOCName', 'RCName', 'FCCName']

g_time_flag = 0


# 获取文件夹下所有log后缀的文件名
def get_file_name():
    filename = []

    for i in glob.glob(r'./*.log'):
        filename.append(i)

    for i in range(len(filename)):
        print(" %d : %s " % (i + 1, filename[i][2:]))

    file_num = input('\n输入文件编号：')

    while not file_num.isdigit() or (int(file_num) - 1) >= len(filename):
        file_num = input('输入错误，请重新输入：')

    return filename[int(file_num) - 1]


# 输出运行时间
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
    def __init__(self, log_name, ex_name):
        self.file_path = log_name
        self.excel_path = './result/' + ex_name + '.xlsx'
        self.log_name = None

    # 获取log数据中对应的模块名
    @staticmethod
    def get_module_name(line):
        if ('ElapsedTime' in line and 'Voltage' in line and 'Current' in line
                and 'RSOC' in line and 'RemCap' in line and 'FullChgCap' in line):

            return ['ElapsedTime', 'Voltage', 'Current', 'RSOC', 'RemCap', 'FullChgCap']

        elif ('~Elapsed(s)' in line and 'Voltage' in line and 'AvgCurrent' in line
              and 'StateofChg' in line and 'RemCap' in line and 'FullChgCap' in line):

            return ['~Elapsed(s)', 'Voltage', 'AvgCurrent', 'StateofChg', 'RemCap', 'FullChgCap']

        return False

    def log_to_excel(self):
        file = open(self.file_path, 'r')
        line = file.readlines()
        file.close()

        try:
            os.mkdir('./result/')
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

        self.log_name = self.get_module_name(line[begin_num])

        if not self.log_name:
            if custom_type:
                self.log_name = custom_name
            else:
                return False

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

        return True

    def print_chart(self):
        file = openpyxl.load_workbook(self.excel_path)
        sheet = file.active
        sheet.freeze_panes = 'A2'

        chart_sheet = file.create_chartsheet('Chart1')

        chart = ScatterChart()

        chart_rsoc = ScatterChart()

        chart.title = 'Project Name Cycle-Test-Curve\n' \
                      '\t\tF/W: **,\tCharge : *V/*A,\tDischarge : *A\n' \
                      '\t\t\t\tTested by: '

        xvalue = Reference(sheet, min_row=2, min_col=sheet.max_column - 5,
                           max_row=sheet.max_row, max_col=sheet.max_column - 5)

        for i in range(sheet.max_column - 4, sheet.max_column + 1):
            yvalue = Reference(sheet, min_row=1, min_col=i,
                               max_row=sheet.max_row, max_col=i)

            series = Series(yvalue, xvalue, title_from_data=True)
            if i == sheet.max_column - 2:
                chart_rsoc.append(value=series)
            else:
                chart.append(value=series)

        chart.x_axis.majorGridlines = None
        chart.y_axis.title = 'Voltage(mV)/Current(mA)/RemCap(mAh)/FullChgCap(mAh)'

        chart_rsoc.y_axis.title = 'RSOC(%)'
        chart_rsoc.y_axis.crosses = 'max'
        chart_rsoc.y_axis.axId = 200
        chart_rsoc.y_axis.majorGridlines = None
        chart_rsoc.x_axis.majorGridlines = None

        chart += chart_rsoc

        chart_sheet.add_chart(chart)

        file.save(self.excel_path)

        file.close()


def main():
    global g_time_flag
    file_name = get_file_name()
    excel_name = input('请输入导出Excel表格文件名（不需要添加后缀）：')

    build_excel = BuildExcel(file_name, excel_name)
    time_count_thread = threading.Thread(target=time_count)
    time_count_thread.daemon = True
    time_count_thread.start()

    print('正在将log数据写入excel，请耐心等待...')
    g_time_flag = 1
    flag = build_excel.log_to_excel()
    g_time_flag = 0

    if flag:
        print('\n写入完成')
    else:
        print('\n不支持该log格式，请参考代码开头自定义数据名')
        return False

    time.sleep(0.1)

    print('正在绘制图表，请耐心等待...')
    g_time_flag = 1
    build_excel.print_chart()
    g_time_flag = 0
    print('\n画图完成，文件保存在result文件夹下')


if __name__ == '__main__':
    main()
    input('按任意键退出')
