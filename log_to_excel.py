import datetime
import glob
import re
import time
import threading
import os
import xlsxwriter
from xml.dom import minidom

# 软件版本 (每次更新后记得修改一下)
tool_version = 'V2.1'

begin_value = 'Sample'  # log数据开头第一个单词，一般为Sample

'''
目前测试可以使用的芯片列表如下：
BQ28Z610, BQ40Z50R2, SN27541， BQ78Z101， BQ20Z45R1, BQ40Z50R1, BQ34Z100
BQ8050, MAX17300
（同时也支持列表上没有芯片，只要log数据中模块名相同即可）
若log数据不支持，在g_module_name中加入相应的模块名即可
'''

# 偏移电流范围
g_drift_current = 100

# 时间显示线程使能
g_time_flag = 0

# 电流倍数
g_current_rate = 1

# 写入参数
g_author = ''
g_chr_voltage = 0
g_term_voltage = 0
g_fw_version = ''
g_project_name = ''
g_project_state = ''

# log数据中的模块名
g_module_name = [
    ['ElapsedTime', '~Elapsed(s)', '~Escape', 'Time'],  # 时间模块名
    ['Voltage', 'VCell (6C:1A)', 'VCell ()'],  # 电压模块名
    ['Current', 'AvgCurrent', 'Current (6C:1C)', 'Current ()'],  # 电流模块名
    ['RSOC', 'StateofChg', 'StateofCharge', 'RepSOC (6C:06)', 'RepSOC ()'],  # RSOC模块名
    ['RemCap', 'RepCap (6C:05)', 'RepCap ()'],  # RC模块名
    ['FullChgCap', 'FullCapRep (6C:10)', 'FullCapRep ()'],  # FCC模块名
    ['Temperature', 'Temp (6C:1B)', 'Temp ()']  # 温度模块名
]

# 芯片型号
g_chip_name = ['sn27541M200', 'bq40z50', 'bq28z610', 'bq34z100']

g_warn_message = []


# 获取文件夹下所有log后缀和csv后缀的文件名
def get_file_name():
    filename = []

    # TI 芯片为log后缀
    for i in glob.glob(r'./*.log'):
        filename.append(i)
    # Maxim 芯片为csv后缀
    for i in glob.glob(r'./*.csv'):
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
    def __init__(self, log_name):
        self.file_path = log_name
        # 根据project name生成Excel文件名
        self.excel_path = ('./result/{0}-Battery-Cycle-Test-Curve-Rev{1}-{2:02}{3:02}{4}.xlsx'
                           .format(g_project_name, g_fw_version, datetime.datetime.now().month,
                                   datetime.datetime.now().day, datetime.datetime.now().year))
        self.log_name = None

        self.chip_name = None

        self.cycle_count = 0
        self.cycle_result = {}

        self.chr_parameter = '0A'
        self.disg_parameter = '0A'

        self.workbook = None

        self.module_num = []

        self.highlight_num = []

    # 获取log数据中对应的模块名
    @staticmethod
    def get_module_name(line):
        global g_module_name
        module_name = []
        for n in range(7):
            for i in g_module_name[n]:
                if i in line:
                    module_name.append(i)
                    break
        if len(module_name) == 7:
            return module_name
        else:
            return False

    # 获取芯片型号
    def get_chip_name(self, line):
        global g_chip_name
        if '.csv' in self.file_path:
            self.chip_name = 'MaximIC'
            return True

        for n in g_chip_name:
            for i in line:
                if re.search(n, i, re.IGNORECASE):
                    self.chip_name = n
                    break
                if re.search('Develop Tool', i, re.IGNORECASE):
                    self.chip_name = 'bq8050'
                    break

            if self.chip_name is not None:
                break

    def log_to_excel(self):
        global begin_value

        file = open(self.file_path, 'r')
        line = file.readlines()
        file.close()

        try:
            os.mkdir('./result/')
        except:
            pass

        # 当为Maxim的芯片时，修改为Time
        if '.csv' in self.file_path:
            begin_value = 'Time'

        # 当获取到begin_value后，定义这一行为开始行
        begin_num = None
        for i in range(len(line)):
            if begin_value in line[i]:
                begin_num = i
                break
            else:
                continue

        # 若没有获取到数据开始行，返回error3报错
        if begin_num is None:
            return 'error3'

        # 获取芯片型号
        self.get_chip_name(line=line[:begin_num])

        # 获取分隔符
        delimiter = ''
        if ',' in line[begin_num]:
            delimiter = ','
        elif '\t' in line[begin_num]:
            delimiter = '\t'
        elif ' ' in line[begin_num]:
            delimiter = ' '

        # 若没有获取到分隔符，返回error2报错
        if not delimiter:
            return 'error2'

        # 获取log数据中个模块的名称
        self.log_name = self.get_module_name(line[begin_num].split(delimiter))

        # 若没有匹配数据，返回error1报错
        if not self.log_name:
            return 'error1'

        i = begin_num
        # 根据分隔符将log数据分隔
        while i < len(line):
            line[i] = line[i].split(delimiter)
            i += 1

        # 生成新的二维数组
        new_line = line[begin_num:]

        self.len_data = len(new_line[1])

        for i in range(len(new_line)):
            # 定义首行数据，获取各数据模块的位置
            if i == 0:
                time_num = new_line[i].index(self.log_name[0])
                voltage_num = new_line[i].index(self.log_name[1])
                current_num = new_line[i].index(self.log_name[2])
                rsoc_num = new_line[i].index(self.log_name[3])
                rc_num = new_line[i].index(self.log_name[4])
                fcc_num = new_line[i].index(self.log_name[5])
                temp_num = new_line[i].index(self.log_name[6])
                new_line[i].extend([' ', 'Time', 'Voltage', 'Current', 'RSOC', 'RC', 'FCC', 'Temperature',
                                    ' ', 'Accumulated', 'Deviation', 'Fuel Gauge Deviation', 'Fuel Gauge Accuracy'])

            elif len(new_line[i]) >= self.len_data:
                # TI芯片下的数据获取
                if self.chip_name != 'MaximIC':
                    # 时间计算
                    if not new_line[i][time_num]:
                        temp_time = ''
                    else:
                        temp_time = round(float(new_line[i][time_num]) / 3600, 6)

                    # 电压计算
                    if not new_line[i][voltage_num]:
                        if i == 1:
                            temp_vol = 0
                        else:
                            temp_vol = ''
                    else:
                        temp_vol = int(new_line[i][voltage_num])

                    # 电流计算
                    if not new_line[i][current_num]:
                        if i == 1:
                            temp_curr = 0
                        else:
                            temp_curr = ''
                    else:
                        temp_curr = abs(int(new_line[i][current_num])) * g_current_rate

                    # RSOC计算
                    if not new_line[i][rsoc_num]:
                        if i == 1:
                            temp_rsoc = 0
                        else:
                            temp_rsoc = ''
                    else:
                        temp_rsoc = int(new_line[i][rsoc_num])

                    # RC计算
                    if not new_line[i][rc_num]:
                        if i == 1:
                            temp_rc = 0
                        else:
                            temp_rc = ''
                    else:
                        temp_rc = int(new_line[i][rc_num])

                    # FCC计算
                    if not new_line[i][fcc_num]:
                        if i == 1:
                            temp_fcc = 0
                        else:
                            temp_fcc = ''
                    else:
                        temp_fcc = int(new_line[i][fcc_num])

                    # 温度计算
                    if not new_line[i][temp_num]:
                        if i == 1:
                            temp_temp = 0
                        else:
                            temp_temp = ''
                    else:
                        temp_temp = float(new_line[i][temp_num])

                # Maxim芯片下的数据获取
                else:
                    # 时间计算
                    try:
                        dt = time.strptime(new_line[i][time_num], '%m/%d/%Y %H:%M:%S')
                    except:
                        dt = time.strptime(new_line[i][time_num], '%m/%d/%Y %H:%M')
                    if i == 1:
                        begin_dt = dt
                    temp_time = round((time.mktime(dt) - time.mktime(begin_dt)) / 3600, 6)

                    # 电压计算
                    if not new_line[i][voltage_num]:
                        temp_vol = ''
                    else:
                        temp_vol = float(new_line[i][voltage_num]) * 1000

                    # 电流计算
                    if not new_line[i][current_num]:
                        temp_curr = ''
                    else:
                        temp_curr = abs(float(new_line[i][current_num])) * g_current_rate

                    # RSOC计算
                    if not new_line[i][rsoc_num]:
                        temp_rsoc = ''
                    else:
                        temp_rsoc = float(new_line[i][rsoc_num])

                    # RC计算
                    if not new_line[i][rc_num]:
                        temp_rc = ''
                    else:
                        temp_rc = float(new_line[i][rc_num])

                    # FCC计算
                    if not new_line[i][fcc_num]:
                        temp_fcc = ''
                    else:
                        temp_fcc = float(new_line[i][fcc_num])

                    # 温度计算
                    if not new_line[i][temp_num]:
                        temp_temp = ''
                    else:
                        temp_temp = float(new_line[i][temp_num])

                ''' 将添加特定数据 '''
                # 部分芯片通讯出现error时，新生成的数据的位置会被打乱，程序可以自动修复
                if re.search('error', new_line[i][-1], re.IGNORECASE) and '~Elapsed(s)' in new_line[0]:
                    pass
                else:
                    new_line[i].extend(' ')

                new_line[i].extend([temp_time, temp_vol, temp_curr, temp_rsoc, temp_rc, temp_fcc, temp_temp])

        # 获取个模块在数组中的位置
        self.module_num = [len(new_line[0]) - 12,  # 时间
                           len(new_line[0]) - 11,  # 电压
                           len(new_line[0]) - 10,  # 电流
                           len(new_line[0]) - 9,  # RSOC
                           len(new_line[0]) - 8,  # RC
                           len(new_line[0]) - 7,  # FCC
                           len(new_line[0]) - 6,  # 温度
                           len(new_line) - 1]  # 数据长度

        # 数据处理算法
        '''
        1. 当通讯错误引起数据读取失败时，程序会根据上下值取中值来估计数据；
        2. 当有一段数据读取失败时，程序会根据前后的数据判断，计算每一个的平均值。
        '''
        for temp_value in self.module_num[1:7]:
            for temp_cow in range(len(new_line)):
                if (len(new_line[temp_cow]) >= len(new_line[1])) and new_line[temp_cow][temp_value] == '':
                    count = 1
                    value_1 = new_line[temp_cow - 1][temp_value]
                    for k in range(temp_cow + 1, len(new_line)):
                        count += 1
                        value_2 = ''
                        if new_line[k][temp_value] != '':
                            value_2 = new_line[k][temp_value]
                            break
                        else:
                            continue
                    if value_2 != '':
                        calibrate_value = value_1 + (value_2 - value_1) / count
                        if self.chip_name != "MaximIC" and (temp_value in [len(new_line[0]) - 11, len(new_line[0]) - 10,
                                                                           len(new_line[0]) - 9, len(new_line[0]) - 8,
                                                                           len(new_line[0]) - 7]):
                            new_line[temp_cow][temp_value] = int(calibrate_value)
                        else:
                            new_line[temp_cow][temp_value] = float(calibrate_value)
                    else:
                        new_line[temp_cow][temp_value] = 0

        # 计算容量
        cap_result = self.cap_accumulated(new_line)

        # 将new_line的数据生成excel
        self.workbook = xlsxwriter.Workbook(self.excel_path)
        worksheet = self.workbook.add_worksheet('data')
        # 冻结首行
        worksheet.freeze_panes(1, 0)
        # 将数据写入Excel
        bg_color = self.workbook.add_format()
        bg_color.set_bg_color('#D8E4BC')

        for n in range(len(new_line)):
            # term点和zero点添加背景色
            if n in self.highlight_num and n != 0:
                for i in range(len(new_line[n])):
                    worksheet.write(n, i, new_line[n][i], bg_color)
            else:
                for i in range(len(new_line[n])):
                    worksheet.write(n, i, new_line[n][i])

        return 'success'

    def cap_accumulated(self, line):
        global g_warn_message

        # 获取新添加的数据的位置
        fcc_num = int(len(line[1])) - 2
        rc_num = fcc_num - 1
        rsoc_num = rc_num - 1
        current_num = rsoc_num - 1
        voltage_num = current_num - 1
        time_num = voltage_num - 1

        # 充放电标志位，chg_flag 和 disg_flag同时等于1时，自动计算容量程序才会开启
        # 即在检查到充电后，下一个放电操作才会计算容量
        chg_flag = 0
        disg_flag = 0

        i = 1
        while i < len(line) and len(line[i]) == len(line[1]):
            global g_term_voltage
            zero_num = 0
            term_num = 0

            ''' 根据漂移电流范围判断充放电 '''
            if not -g_drift_current < line[i][current_num] < g_drift_current:
                begin_num = i
                end_num = 0
                while i < len(line) and len(line[i]) == len(line[1]):
                    end_num = i
                    if -g_drift_current < line[i][current_num] < g_drift_current:
                        break
                    else:
                        i += 1

                ''' 若充放电时间过短，跳过该充放电阶段 '''
                if (end_num - begin_num) <= 10 or end_num == 0:
                    continue

                ''' 通过RSOC大小判断充放电 '''
                if line[begin_num][rsoc_num] < line[end_num][rsoc_num]:
                    # 充电
                    chg_flag = 1

                    # 判断充电电流大小
                    chg_curr = line[round((end_num - begin_num) / 10) + begin_num][current_num]
                    self.chr_parameter = str(round(chg_curr / 100) / 10) + 'A'
                elif line[begin_num][rsoc_num] > line[end_num][rsoc_num]:
                    # 放电
                    disg_flag = 1

                    # 判断恒流放电还是恒功率放电
                    disg_curr_mid = line[round((end_num - begin_num) / 2) + begin_num][current_num]

                    # 取三个时间点的放电电流，若三个时间点之间的误差不超过500，则为恒流放电
                    if abs(disg_curr_mid - line[begin_num + 5][current_num]) < 500 and abs(
                            disg_curr_mid - line[end_num - 5][current_num]) < 500 and abs(
                        line[begin_num + 5][current_num] - line[end_num - 5][current_num]) < 500:
                        self.disg_parameter = str(round(disg_curr_mid / 100) / 10) + 'A'

                    else:
                        # 恒功率放电时，取三个时间点的功率计算平均值
                        disg_power = round((disg_curr_mid * line[round((end_num - begin_num) / 2)][voltage_num] +
                                            line[begin_num + 5][current_num] * line[begin_num + 5][voltage_num] +
                                            line[end_num - 5][current_num] * line[end_num - 5][voltage_num]) / 3000000)
                        self.disg_parameter = str(int(disg_power / 10) * 10) + 'W'

                    # 判断为放电阶段时，记录下rc 0点和term点，用于计算容量差
                    for n in range(begin_num, end_num):
                        if line[n][rsoc_num] == 0 and zero_num == 0:
                            zero_num = n
                        if line[n][voltage_num] < g_term_voltage and term_num == 0:
                            term_num = n

                if chg_flag == 1 and disg_flag == 1:

                    self.cycle_count += 1

                    line[begin_num - 1].extend([' ', 0])
                    for n in range(begin_num, end_num + 1):
                        # 容量计算公式
                        temp_cap = ((line[n][time_num] - line[n - 1][time_num]) *
                                    (line[n][current_num] + line[n - 1][current_num]) / 2 + line[n - 1][-1])

                        line[n].extend([' ', temp_cap])

                    ''' Maxim芯片zero点计算方法 '''
                    if self.chip_name == 'MaximIC':
                        t_2 = 0
                        t_1 = 0
                        t_0 = 0

                        for n in range(begin_num, end_num + 1):
                            if 2 < line[n][rsoc_num] <= 3:
                                t_2 += 1
                            if 1 < line[n][rsoc_num] <= 2:
                                t_1 += 1
                            if 0 <= line[n][rsoc_num] <= 1:
                                t_0 += 1
                        if t_0 < t_1 and t_2 < (t_1 + t_0) and t_0 != 0:
                            zero_num = end_num - (t_1 + t_0 - 2 * t_2)
                        elif t_0 >= t_1:
                            zero_num = end_num - (t_0 - t_1)
                        pass

                    ''' Maxim芯片term点确定方法 '''
                    if self.chip_name == 'MaximIC':
                        try:
                            # MAX17300
                            fstat_num = line[0].index('FStat (6C:3D)')
                        except:
                            # MAX17201
                            fstat_num = line[0].index('FStat ()')
                        for n in range(begin_num, end_num + 1):
                            try:
                                if int(line[n][fstat_num], 16) & 0x100:
                                    term_num = n
                                    break
                            except:
                                pass

                    ''' 自动确定term点方式 '''
                    # 检测 GaugeStat 中的 EDV 位，若EDV位为1，则当该时刻为term点
                    if self.chip_name != 'MaximIC' and 'GaugeStat' in line[0]:
                        gauge_status_num = line[0].index('GaugeStat')

                        for n in range(begin_num, end_num + 1):
                            try:
                                if int(line[n][gauge_status_num], 16) & 0x20:
                                    term_num = n
                                    break
                            except:
                                pass

                    ''' Term点未出现情况 '''
                    # 当term点没有出现，在end_num往后20个时间点内检测是否出现term点，若出现则定为新的term点
                    if term_num == 0:
                        i = end_num
                        while i - end_num < 20 and i < len(line):
                            # bq40z50的term点计算
                            if self.chip_name != 'MaximIC' and 'GaugeStat' in line[0]:
                                if (len(line[i]) >= self.len_data and
                                        line[i][gauge_status_num] and
                                        int(line[i][gauge_status_num], 16) & 0x20):
                                    term_num = i
                                    break
                            elif self.chip_name == 'MaximIC':
                                if (len(line[i]) >= self.len_data and
                                        line[i][fstat_num] and
                                        int(line[i][fstat_num], 16) & 0x100):
                                    term_num = i
                                    break
                            # 一般情况下根据term_voltage计算
                            elif line[i][voltage_num] < g_term_voltage:
                                term_num = i
                                break
                            i += 1

                    if term_num != 0:
                        # 计算放电结束点到新term点的容量
                        for n in range(end_num + 1, term_num + 1):
                            temp_cap = ((line[n][time_num] - line[n - 1][time_num]) *
                                        (line[n][current_num] + line[n - 1][current_num]) / 2 + line[n - 1][-1])
                            line[n].extend([' ', temp_cap])
                    else:
                        # 若还是没有发现新term点，把end_num点代入term点计算，并提示warning信息
                        term_num = end_num
                        g_warn_message.append('Cycle{0} 未发现Term点，代入放电最后一个时刻点进行计算。'.format(self.cycle_count))

                    ''' 一般情况下计算 '''
                    if zero_num != 0 and term_num != 0:
                        cap_dev = line[zero_num][-1] - line[term_num][-1]
                        cap_dev_percentage = cap_dev / line[term_num][-1]
                    else:
                        cap_dev = None
                        cap_dev_percentage = None

                    ''' 特殊情况1 '''
                    # 当RSOC瞬间跳为0时，检测上一时刻是否为1，若不是，cap_dev_percentage就为两个时刻的rsoc的差值
                    if cap_dev_percentage == 0:
                        if line[zero_num - 1][rsoc_num] > 1:
                            cap_dev_percentage = (line[zero_num][rsoc_num] - line[zero_num - 1][rsoc_num]) / 100

                    ''' 特殊情况2 '''
                    # 当rsoc没有出现0的情况下，cap_dev_percentage为term点的rsoc值
                    if zero_num == 0:
                        cap_dev = 0
                        cap_dev_percentage = line[term_num][rsoc_num] / 100

                    if line[begin_num][fcc_num] != 0:
                        cap_percentage = line[term_num][-1] / line[begin_num][fcc_num]
                    else:
                        cap_percentage = 0

                    if cap_percentage > 1:
                        cap_percentage = 1 / cap_percentage
                    line[term_num].extend(
                        [cap_dev, '{:.2%}'.format(cap_dev_percentage), '{:.2%}'.format(cap_percentage)])

                    if -0.06 <= cap_dev_percentage <= 0.06:
                        self.cycle_result['Cycle ' + str(self.cycle_count)] = ('{:.2%} PASS'.format(cap_dev_percentage))
                    else:
                        self.cycle_result['Cycle ' + str(self.cycle_count)] = ('{:.2%} FAIL'.format(cap_dev_percentage))

                    self.highlight_num.extend([term_num, zero_num])
                    chg_flag = 0

            disg_flag = 0
            i += 1

        return True

    def print_chart(self):
        # 创建图表
        chartsheet = self.workbook.add_chartsheet('chart')
        chart = self.workbook.add_chart({'type': 'scatter',
                                         'subtype': 'straight'})

        # cycle test结果输出
        result_title = ''
        for i in self.cycle_result:
            result_title += '{0} : {1}   '.format(i, self.cycle_result[i])

        # 设置图表标题
        chart.set_title({'name': '{0} Battery Pack Cycle-Test-Curve\n'
                                 'State: Proto {1},   FW: {2},   Charge : {3}V/{4},   Discharge : {5}\n'
                                 '{6}\n'
                                 '\t\t\t\t\tTested by:{7}'
                        .format(g_project_name, g_project_state, g_fw_version, g_chr_voltage,
                                self.chr_parameter, self.disg_parameter, result_title, g_author)})

        # 设置主Y轴标题
        chart.set_y_axis({'name': 'Voltage(mV)/Current(mA)/RemCap(mAh)/FullChgCap(mAh)'})
        # 设置副Y轴标题
        chart.set_y2_axis({'name': 'RSOC(%)/Temperature(\'C)'})

        ''' 添加系列数据 '''
        # 电压
        chart.add_series({'name': ['data', 0, self.module_num[1]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[1], self.module_num[7], self.module_num[1]], })
        # 电流
        chart.add_series({'name': ['data', 0, self.module_num[2]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[2], self.module_num[7], self.module_num[2]], })
        # RSOC
        chart.add_series({'name': ['data', 0, self.module_num[3]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[3], self.module_num[7], self.module_num[3]],
                          'y2_axis': 1, })
        # RC
        chart.add_series({'name': ['data', 0, self.module_num[4]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[4], self.module_num[7], self.module_num[4]], })
        # FCC
        chart.add_series({'name': ['data', 0, self.module_num[5]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[5], self.module_num[7], self.module_num[5]], })
        # 温度
        chart.add_series({'name': ['data', 0, self.module_num[6]],
                          'categories': ['data', 1, self.module_num[0], self.module_num[7], self.module_num[0]],
                          'values': ['data', 1, self.module_num[6], self.module_num[7], self.module_num[6]],
                          'y2_axis': 1, })

        # 写入图表
        chartsheet.set_chart(chart)

        self.workbook.close()


def main():
    global g_time_flag
    global g_term_voltage
    global g_author
    global g_chr_voltage
    global g_fw_version
    global g_project_name
    global g_project_state
    global g_warn_message
    global begin_value
    global g_drift_current
    global g_current_rate

    # 从config.xml文件中读取配置
    try:
        config = minidom.parse('./config.xml')
        begin_value = config.getElementsByTagName('begin_value')[0].firstChild.data
        g_drift_current = int(config.getElementsByTagName('drift_current')[0].firstChild.data)
        g_current_rate = int(config.getElementsByTagName('current_rate')[0].firstChild.data)
    except:
        pass

    print("######## 煲机数据自动处理工具" + tool_version + " ########")
    print('开始标记 : {0} (Maxim 芯片可忽略这一项)\n漂移电流范围 : {1} mA\t电流倍数 : {2}\n'
          .format(begin_value, g_drift_current, g_current_rate))

    file_name = get_file_name()
    g_project_name = input('请输入项目名称：')
    g_author = input('请输入作者：')
    g_project_state = input('请输入Proto阶段：')
    g_fw_version = input('请输入软件版本：')
    g_chr_voltage = input('请输入充电电压 (V)：')
    g_term_voltage = int(input('输入term_voltage (mV)：'))

    build_excel = BuildExcel(file_name)
    time_count_thread = threading.Thread(target=time_count)
    time_count_thread.daemon = True
    time_count_thread.start()

    print('正在将log数据写入excel，请耐心等待...')
    g_time_flag = 1
    flag = build_excel.log_to_excel()
    g_time_flag = 0

    # 结果分析
    if flag == 'success':
        print('\n写入完成')
    elif flag == 'error1':
        print('\n暂时不支持该log格式')
        return False
    elif flag == 'error2':
        print('\n未知log数据分隔符')
        return False
    elif flag == 'error3':
        print('\n未获取到开始行，请检查log数据开始行是否以 %s 开头' % begin_value)
        return False

    # 打印警报信息
    if g_warn_message:
        print('\nWarning:')
        for i in g_warn_message:
            print('\t%s' % i)
        print('\tTerm voltage 为 {0} mV，请确认是否有误。若输入错误，请重新执行程序。\n'.format(g_term_voltage))

    time.sleep(0.1)

    print('正在绘制图表，请耐心等待...')
    g_time_flag = 1
    build_excel.print_chart()
    g_time_flag = 0
    print('\n画图完成，文件保存在result文件夹下')


if __name__ == '__main__':
    main()
    input('按任意键退出')
