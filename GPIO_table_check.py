# coding=utf-8
# 2018.7.3 从文件中提取所需数据
# python 3.7

import os
import re
import xlwings as xw
import copy
from math import ceil
import xlsxwriter


# Class Groups


# 读取 pstxnet.dat，pstxprt.dat，pstchip.dat 及 EXP文件的模块
class ExtractIOData:
    """对DSN导出的报告进行数据处理"""

    def __init__(self):
        root_path = os.getcwd()
        self._root_path = os.path.join('\\'.join(root_path.split('\\')[:-1]), 'input')
        self._output_path = os.path.join('\\'.join(root_path.split('\\')[:-1]), 'output')
        self._output_excel_path = os.path.join(self._output_path, 'Initial_GPIO_Table.xlsx')

        # pstxnet.dat
        self.all_net_list_ = []
        self.net_node_dict_ = {}
        self.net_node_list_ = []

        # pstxprt.dat
        self.all_node_list_ = []
        self.node_page_dict_ = {}
        self.all_res_dict_ = {}
        self.all_res_list_ = []
        self.ic_ext_icname_dict_ = {}

        # pstchip.dat
        self.pin_name_list_ = []
        self.ext_icname_pin_num_dict_ = {}

        # exp
        self.ic_ni_dict_ = {}

    def extract_pstxnet(self):
        """提取pstxnet.dat的数据"""
        try:
            with open(os.path.join(self._root_path, 'pstxnet.dat'), 'r') as file1:
                content1 = file1.read().split('NET_NAME')
                for ind1 in range(len(content1)):
                    content1[ind1] = content1[ind1].split('\n')
                for x in content1:
                    node_list = []
                    self.all_net_list_.append(x[1][1:-1])
                    for y_idx in range(len(x)):
                        if x[y_idx].find('NODE_NAME') > -1:
                            node_list.append([x[y_idx].split('NODE_NAME\t')[-1].split(' ')[0], x[y_idx + 2].split("'")[1]])
                    node_flatten_list = list(flatten([[x[1][1:-1]] + node_list]))
                    # print('node_flatten_list', node_flatten_list)
                    self.net_node_dict_[x[1][1:-1]] = node_flatten_list
                    self.net_node_list_.append(node_flatten_list)
                self.all_net_list_ = self.all_net_list_[1:]

            return self.all_net_list_, self.net_node_list_, self.net_node_dict_
        except FileNotFoundError:
            error_message = 'Missing pstxnet.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_pstxprt(self):
        """提取pstxprt.dat的数据"""
        try:
            with open(os.path.join(self._root_path, 'pstxprt.dat'), 'r') as file2:
                content2 = file2.read().split('PART_NAME')

                primitive_list = []

                for ind2 in range(len(content2)):
                    content2[ind2] = content2[ind2].split('\n')

                for x in content2:
                    # print(x)
                    # print('\n')
                    node = x[1].split(' ')[1]
                    self.all_node_list_.append(node)

                    pattern = re.compile(r".*?_.*?_(.*?)_.*?")
                    res_val = pattern.findall(x[1].split(' ')[2])

                    if x[0] == '':
                        pattern2 = re.compile(r".*?'(.*?)'.*?")
                        pattern3 = re.compile(r".*?@.*?@(.*?)\..*?")
                        node1 = x[1].split(' ')[1]
                        self.ic_ext_icname_dict_[node1] = pattern2.findall(x[1])[0]

                        if pattern3.findall(x[5])[0].upper() == 'RESISTOR':
                            self.all_res_list_.append(node1)

                    if res_val:
                        self.all_res_dict_[node] = res_val[0]

                    if x[0] == '':
                        primitive_list.append(x[1].split("\'")[1])
                    if 'page' in x[6]:
                        page_now = x[6].split(':')[-1].split('_')[0]
                    else:
                        page_now = x[7].split(':')[-1].split('_')[0]
                    if self.node_page_dict_.get(page_now):
                        self.node_page_dict_[page_now] += [self.all_node_list_[-1]]
                    else:
                        self.node_page_dict_[page_now] = [self.all_node_list_[-1]]

            return self.all_node_list_, self.node_page_dict_, self.all_res_dict_, \
                self.all_res_list_, self.ic_ext_icname_dict_
        except FileNotFoundError:
            error_message = 'Missing pstxprt.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_pstchip(self, GPIO_pin_name_list_org=None, func=None):
        try:
            with open(os.path.join(self._root_path, 'pstchip.dat'), 'r') as file5:

                # 用于测试阶段， 对表中的pin name进行变形
                if func == 'reshape_pin_name':
                    content5 = file5.read().split('primitive')
                    pattern1 = re.compile(r".*?'(.*?)':")
                    self.pin_name_list_ = []
                    for y in GPIO_pin_name_list_org:
                        y_flag = True
                        # 先对输入pin进行处理
                        # 去除前后空格
                        y = y.strip()
                        if y.upper().find('GROUP') > -1:
                            self.pin_name_list_.append(y)
                        else:
                            # 对有无/符号进行分类
                            if y.find('/') > -1:
                                y = '/'.join(map(lambda x: x.strip(), y.split('/')))
                                y = y[:31]
                                # print(y)
                            for x in content5:
                                key_items = pattern1.findall(x)
                                for key_idx in range(len(key_items)):

                                    if key_items[key_idx] == y:
                                        self.pin_name_list_.append(y)
                                        # print(1, y)
                                        y_flag = False
                                        break
                                    elif key_items[key_idx].find(y + '/') > -1:
                                        # print(2, key_items[key_idx])
                                        self.pin_name_list_.append(key_items[key_idx])
                                        y_flag = False
                                        break
                                if y_flag is False:
                                    break
                            # pin名错误
                            if y_flag:
                                self.pin_name_list_.append(None)
                    return self.pin_name_list_

                # 用于第一阶段, 获取pin name
                if func == 'get_pin_name':
                    content5 = file5.read().split('primitive')
                    pattern1 = re.compile(r".*?'(.*?)':")
                    self.pin_name_list_ = []
                    for x in content5:
                        key_items = pattern1.findall(x)
                        for key_idx in range(len(key_items)):

                            if str(key_items[key_idx]).startswith('GPP_'):
                                self.pin_name_list_.append(key_items[key_idx])
                    return sorted(self.pin_name_list_)

                # 获取ic及其pin的数量信息
                if func == 'get_ic_pin_number':
                    content = file5.read().split('end_primitive')
                    pattern = re.compile(r".*?primitive '(.*?)'")

                    for c_item in content:
                        key_item = pattern.findall(c_item)
                        if key_item:
                            self.ext_icname_pin_num_dict_[key_item[0]] = c_item.count('PIN_NUMBER')

                    return self.ext_icname_pin_num_dict_

                # 获取pin name对应的pin location
                if func == 'get_pin_name_pin_location':
                    content5 = file5.read().split('primitive')
                    pattern1 = re.compile(r".*?'(.*?)':")
                    key_list = []
                    value_list = []
                    pattern2 = re.compile(r".*?PIN_NUMBER='\((.*?)\)';")
                    for x in content5:
                        pattern1_data = pattern1.findall(x)
                        if pattern1_data:
                            key_list.append(pattern1_data)
                        pattern2_data = pattern2.findall(x)
                        if pattern2_data:
                            value_half0_list = []
                            for pattern2_data_item in pattern2_data:

                                # 为了处理BA45,0,0,0,0,0,0,0,0,0,0,0,0这种情况
                                if pattern2_data_item.find(',') > -1:
                                    value_half1_list = list(set(pattern2_data_item.split(',')))
                                    try:
                                        value_half1_list.remove('0')
                                        value_half0_list.append(value_half1_list[0])
                                    # 二极管，三极管也有逗号，所以要区分开
                                    except:
                                        value_half0_list.append(pattern2_data_item)

                                else:
                                    value_half0_list.append(pattern2_data_item)
                            value_list.append(value_half0_list)

                    pin_name_pin_location_dict = dict(zip(flatten(key_list), flatten(value_list)))
                    return pin_name_pin_location_dict
        except FileNotFoundError:
            error_message = 'Missing pstchip.dat file'
            create_error_message(self._output_excel_path, error_message)
            raise FileNotFoundError

    def extract_exp(self, file_name):
        with open(os.path.join(self._root_path, file_name), 'r', encoding='gb18030') as file7:
            file7.readline()
            topic_list = file7.readline().split('\t')
            # ADD:添加异常处理机制
            bom_idx = topic_list.index('"BOM"')

            # try:
            for line in file7.readlines():
                line_list = line.split('\t')
                line_id = line_list[1][1:-1]
                # print(line_id)
                line_ni = line_list[bom_idx][1:-1]
                self.ic_ni_dict_[line_id.upper()] = line_ni
        # print(self.ic_ni_dict_)
        return self.ic_ni_dict_


# 读取表中的列项并跑出pin所连接的信息
class ExtractPinData:

    def __init__(self, sheet):

        self.sheet = sheet

        # 列项信息
        self.pin_name_coord = None
        self.pin_location_coord = None

        # 列项详细信息
        self.GPIO_pin_name_list = []
        self.GPIO_pin_location_list = []

        # 详细走线信息
        self.pin_net_node_list = []
        self.pin_net_node_dict = {}
        self.error_all_list = []

    def get_headline_detail_info(self, pin_name_coord=None, pin_location_coord=None):
        """获取列项的详细信息"""
        # 取得每一个需要check的列的数据并存成list的形式
        group_flag = True
        # 获取要check的sheet中的列项信息
        while group_flag:
            col_idx = len(self.sheet.range(pin_name_coord).options(expand='table').value) + pin_name_coord[0] - 1
            pin_name_list = [str(x) for x in self.sheet.range((pin_name_coord[0], pin_name_coord[1]),
                                                              (col_idx, pin_name_coord[1])).value]

            pin_location_list = [str(x) for x in self.sheet.range((pin_name_coord[0], pin_location_coord[1]),
                                                                  (col_idx, pin_location_coord[1])).value]

            self.GPIO_pin_name_list.append(pin_name_list + ['Group'])
            self.GPIO_pin_location_list.append(pin_location_list + ['Group'])

            col_idx += 1
            # 不为none说明下面还有group
            if self.sheet.range(col_idx + 1, pin_name_coord[1]).value:
                pin_name_coord = (col_idx + 1, pin_name_coord[1])
                pin_location_coord = (col_idx + 1, pin_location_coord[1])
            else:
                group_flag = False

        return self.GPIO_pin_name_list,  self.GPIO_pin_location_list

    def get_detail_layout_info(self, net_node_list, GPIO_pin_name_list_org,
                               all_res_list, IC_pin_num_dict, Exclude_Net_List, ic_ni_dict):
        """获取详细走线信息"""
        net_node_copy_list = copy.deepcopy(net_node_list)
        # 对每一个pin name进行数据处理，找出与之对应的走线规则
        for pin_idx in range(len(GPIO_pin_name_list_org)):

            # 遍历pin name去除分割行
            if str(GPIO_pin_name_list_org[pin_idx]).upper().find('GROUP') == -1:
                # 遍历没有拼写错误的所有pin脚
                if GPIO_pin_name_list_org[pin_idx]:
                    node_item_flag = False
                    # final_flag = False
                    no_nc_flag = False
                    for node_item in net_node_list:
                        net_item = node_item[0]
                        # 找到pin脚连接的信号线信息
                        if GPIO_pin_name_list_org[pin_idx] in node_item[1:]:
                            pin_net_node_list1 = [net_item]
                            node_item1 = copy.deepcopy(node_item)
                            node_item1.pop(node_item1.index(GPIO_pin_name_list_org[pin_idx], 1) - 1)
                            node_item1.pop(node_item1.index(GPIO_pin_name_list_org[pin_idx], 1))
                            node_item1 = node_item1[1::2]
                            flagfour = True
                            split_flag = True
                            split_node_list = []
                            layer_num = 0
                            layer_add_num_dict = {}

                            # 如果匹配到NC，因为NC在最后，所以匹配到NC说明前面都不匹配
                            if pin_net_node_list1[0] == 'NC':

                                # 如果只匹配到NC
                                if no_nc_flag is False:
                                    # if str(GPIO_pupd_list_org[pin_idx]) != 'None':
                                    pin_net_node_list1 = []
                                    # print('NC', pin_net_node_list1)
                                    self.pin_net_node_dict[GPIO_pin_name_list_org[pin_idx]] = pin_net_node_list1
                                    node_item_flag = True
                                # 如果除了NC还匹配到其他
                                # 这段代码好像没用，但是我也不想动它了
                                # else:
                                #     final_flag = True
                            else:
                                no_nc_flag = True
                                # final_flag = False
                                while flagfour:
                                    node_item3 = []

                                    break_flag = False
                                    split_out_flag = False
                                    # if pin_name_list[pin_idx] == 'GPIO1':
                                    #     print(1, node_item1)
                                    if split_flag:
                                        split_node_list.append(copy.deepcopy(node_item1))
                                    for x_idx in range(len(node_item1)):
                                        # IC_flag = False
                                        next_flag = False
                                        all_break = False
                                        add_num = 0
                                        layer_num += 1

                                        item1 = node_item1[x_idx]
                                        # print(2, item1)
                                        split_node_list[-1].pop(split_node_list[-1].index(node_item1[x_idx]))

                                        # 对下一个经过的元器件是否是NI进行判断
                                        if ic_ni_dict[item1] == 'NI':
                                            pin_net_node_list1.append('NI')
                                            split_node_list.append([])
                                            split_flag = False
                                            add_num += 1
                                        # 如果上电
                                        else:
                                            # 大于4并且不是排阻则说明到另外一个芯片了，停止
                                            if item1 not in all_res_list and IC_pin_num_dict[item1] >= 3:
                                                split_flag = False
                                                pin_net_node_list1.append(item1)
                                                add_num += 1
                                                split_node_list.append([])
                                            else:
                                                # print(pin_net_node_list1)
                                                # 判断是否为终止端元器件（中间的元器件会出现两次）
                                                if node_item1.count(item1) >= 1:
                                                    # count = 0
                                                    for node_item2 in net_node_copy_list:
                                                        # count += 1
                                                        add_sch_flag = False
                                                        # 找到元器件所连接的另一根线
                                                        # 如果这次经过的线与上次或第一次相同，则退出
                                                        if item1 in node_item2 and node_item2 != node_item \
                                                                and node_item2 != node_item3:

                                                            # 如果中间没有经过过这个元器件则进入循环
                                                            if node_item2[0] not in pin_net_node_list1:
                                                                add_sch_flag = True
                                                                pin_net_node_list1.append(item1)
                                                                pin_net_node_list1.append(node_item2[0])
                                                                add_num += 2
                                                                if node_item2[0] != node_item[0] and node_item2 \
                                                                        != node_item3:
                                                                    if node_item2[0] in Exclude_Net_List:
                                                                        split_node_list.append([])
                                                                        split_flag = False
                                                                        break

                                                                    node_item1 = copy.deepcopy(node_item2)
                                                                    node_item3 = copy.deepcopy(node_item2)

                                                                    node_item1.pop(node_item1.index(item1) - 1)
                                                                    node_item1.pop(node_item1.index(item1))

                                                                    node_item1 = node_item1[1::2]

                                                                    next_flag = True
                                                                    split_flag = True
                                                                    break_flag = True
                                                                    break
                                                                    # break 不要break，是因为可能元器件有超过两个pin，
                                                                    # 要所有都遍历到, 虽然速度会变慢
                                                                else:
                                                                    all_break = True

                                                        if net_node_copy_list[
                                                            -1] == node_item2 and add_sch_flag is \
                                                                False:
                                                            split_flag = False
                                                            add_num += 1
                                                            split_node_list.append([])
                                                            pin_net_node_list1.append(item1)

                                                            if pin_net_node_list1 not in self.pin_net_node_list:
                                                                self.pin_net_node_list.append(pin_net_node_list1)

                                                            if self.pin_net_node_dict.get(
                                                                    GPIO_pin_name_list_org[pin_idx]):
                                                                pin_net_dict_list = self.pin_net_node_dict[
                                                                    GPIO_pin_name_list_org[pin_idx]]
                                                                pin_net_dict_list.append(pin_net_node_list1)
                                                                self.pin_net_node_dict[GPIO_pin_name_list_org[
                                                                    pin_idx]] = pin_net_dict_list
                                                            else:
                                                                self.pin_net_node_dict[GPIO_pin_name_list_org[
                                                                    pin_idx]] = [pin_net_node_list1]

                                        if all_break:
                                            pass
                                        else:
                                            layer_add_num_dict[layer_num] = add_num
                                            split_node_flag = True
                                            before_layer_num = 0
                                            if next_flag is False:
                                                if split_flag is False or node_item1[-1] == item1:
                                                    split_flag = True
                                                    split_node_list.pop(-1)
                                                    before_layer_num = copy.deepcopy(layer_num)
                                                    layer_num = len(split_node_list)
                                                    if split_node_list:
                                                        try:
                                                            while not split_node_list[-1]:
                                                                split_node_list.pop(-1)
                                                                layer_num -= 1
                                                                # print('layer_num', layer_num)
                                                        except IndexError:
                                                            pass

                                                    if split_node_list:
                                                        node_item1 = split_node_list[-1]
                                                        # if len(node_item1) == 1:
                                                        split_node_list.pop(-1)
                                                        # if node_item1[-1] == item1:
                                                        break_flag = True
                                                        split_node_flag = True

                                                        # print('split_node_flag', split_node_flag)
                                                    else:
                                                        split_node_flag = False
                                                        flagfour = False
                                                        if node_item1[-1] == item1:
                                                            split_out_flag = True

                                                    if pin_net_node_list1 not in self.pin_net_node_list:
                                                        self.pin_net_node_list.append(pin_net_node_list1)

                                                    if 'NI' in pin_net_node_list1:
                                                        if self.pin_net_node_dict.get(GPIO_pin_name_list_org[pin_idx]):
                                                            pass
                                                        else:
                                                            self.pin_net_node_dict[
                                                                GPIO_pin_name_list_org[pin_idx]] = []
                                                    else:
                                                        if self.pin_net_node_dict.get(GPIO_pin_name_list_org[pin_idx]):
                                                            pin_net_dict_list = self.pin_net_node_dict[
                                                                GPIO_pin_name_list_org[pin_idx]]
                                                            pin_net_dict_list.append(pin_net_node_list1)
                                                            self.pin_net_node_dict[GPIO_pin_name_list_org[pin_idx]] = \
                                                                pin_net_dict_list
                                                        else:
                                                            self.pin_net_node_dict[GPIO_pin_name_list_org[pin_idx]] = \
                                                                [pin_net_node_list1]

                                            if split_node_flag:
                                                if before_layer_num:
                                                    for layer_idx in range(layer_num, before_layer_num + 1):
                                                        # print(layer_idx)
                                                        if layer_add_num_dict[layer_idx] != 0:
                                                            pin_net_node_list1 = pin_net_node_list1[:
                                                                                                    -
                                                                                                    layer_add_num_dict[
                                                                                                        layer_idx]]

                                                        layer_add_num_dict.pop(layer_idx)
                                                    layer_num -= 1
                                            if break_flag:
                                                break

                                            if split_out_flag:
                                                break

                            # if node_item == net_node_list[-1] and final_flag:
                            #     node_item_flag = True
                            if node_item_flag:
                                break

                else:
                    self.error_all_list.append(pin_idx)
                    self.pin_net_node_dict[GPIO_pin_name_list_org[pin_idx]] = [['misspelled']]
        return self.pin_net_node_list, self.pin_net_node_dict, self.error_all_list


# Function Groups


# 将多维list展开成一维
def flatten(a):
    if not isinstance(a, (list,)) and not isinstance(a, (tuple,)):
        return [a]
    else:
        b = []
        for item in a:
            b += flatten(item)
    return b


# 将信号线分为电源线与地线
def get_exclude_netlist(netlist):  # netlist = All_Net_List
    # Get pwr and gnd net list

    PWR_Net_KeyWord_List = ['^\+.*', '^-.*',
                            'VREF|PWR|VPP|VSS|VREG|VCORE|VCC|VT|VDD|VLED|PWM|VDIMM|VGT|VIN|[^S](VID)|VR',
                            'VOUT|VGG|VGPS|VNN|VOL|VSD|VSYS|VCM|VSA',
                            '.*\+[0-9]V.*', '.*\+[0-9][0-9]V.*']
    PWR_Net_List = [net for net in netlist for keyword in PWR_Net_KeyWord_List if re.findall(keyword, net) != []]
    PWR_Net_List = sorted(list(set(PWR_Net_List)))

    GND_Net_List = [net for net in netlist if net.find('GND') > -1]
    GND_Net_List = sorted(list(set(GND_Net_List)))

    # 被排除的线：地线和电源线
    Exclude_Net_List = sorted(list(set(PWR_Net_List + GND_Net_List)))

    return Exclude_Net_List, PWR_Net_List, GND_Net_List


# 定义错误输出
def create_error_message(excel_path, error_message):
    # 創建excel
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet('error_message')

    error_format = workbook.add_format({'font_size': 22})

    worksheet.write('A1', 'Program running error:', error_format)
    worksheet.write('B2', error_message + ', please check and try again!', error_format)

    workbook.close()


# 生成完整版 GPIO TABLE
def generate_report():
    """生成报告"""
    # 生成读取数据的类
    extractIOData = ExtractIOData()
    input_path = extractIOData._root_path
    output_path = extractIOData._output_path
    output_excel_path = extractIOData._output_excel_path
    excel_exist_flag = False

    # 生成结果表格的绝对路径
    input_excel_path = os.path.join(input_path, 'Initial_GPIO_Table.xlsx')

    # 如果有初始表格存在说明是更新版本，所以直接在表格中获取PIN的信息
    if os.path.exists(input_excel_path):

        excel_exist_flag = True

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(input_excel_path)
        sht = wb.sheets[0]

        # 表格中所有信息数组
        content = sht.range((2, 1)).options(expand='table').value

        # 获取列项信息
        extractPinData = ExtractPinData(sht)
        pin_name_coord, pin_location_coord = (2, 1), (2, 2)

        # 取得每一个需要check的列的数据并存成list的形式
        GPP_pin_name_list, GPP_pin_location_list = \
            extractPinData.get_headline_detail_info(pin_name_coord, pin_location_coord)
        # 变为一维list
        GPP_pin_name_list = list(flatten(GPP_pin_name_list))
        GPP_pin_location_list = list(flatten(GPP_pin_location_list))

        GPP_pin_name_list = GPP_pin_name_list[1: -1]
        GPP_pin_location_list = GPP_pin_location_list[1: -1]

        row_length = len(content)
        column_length = len(content[0])

        rng = xw.Range((2, 1), (row_length + 1, column_length)).columns
        # rng = xw.Range((2, 1), (229, 10))

        columns_list = []
        for rng_idx in range(len(rng)):
            columns_list.append(rng[rng_idx].value)

        wb.close()
        app.quit()

        os.remove(input_excel_path)

    # 初始没有表格存在说明是第一个版本，则需生成 PIN_NAME, PIN_LOCATION
    else:
        # ************************生成 pin name list 和 pin location list **********************************

        # 生成所有带GPP的pin name
        pin_name_list = extractIOData.extract_pstchip(func='get_pin_name')
        # 对pin name进行排序
        GPP_pin_name_list = []

        # 获取pin name抬头，例如：GPP_A
        GPP_title_word_list = [item.split('/')[0][:5] for item in pin_name_list]
        GPP_title_word_list = sorted(list(set(GPP_title_word_list)))

        # 生成pin name list
        # 根据开头找出所有符合规律的pin name并按字母数字排序
        for GPP_word in GPP_title_word_list:
            GPP_title_seg_num_dict = {}
            GPP_title_seg_num_list = []
            # 对pin name分段处理，分group
            GPP_title_seg_list = [pin_item for pin_item in pin_name_list if pin_item.startswith(GPP_word)]
            for GPP_title_seg_item in GPP_title_seg_list:
                GPP_title_seg_split_item = GPP_title_seg_item.split('/')
                try:
                    GPP_title_seg_num_list.append(int(GPP_title_seg_split_item[0][5:]))
                    GPP_title_seg_num_dict[GPP_title_seg_split_item[0][5:]] = GPP_title_seg_item
                # 特殊情况是不定的
                except:
                    # 数字前有下划线的情况
                    try:
                        GPP_title_seg_num_list.append(int(GPP_title_seg_split_item[0][6:]))
                        GPP_title_seg_num_dict[GPP_title_seg_split_item[0][6:]] = GPP_title_seg_item
                    # 数字后有下划线的情况
                    except:
                        # print(GPP_title_seg_split_item)
                        GPP_title_seg_num_list.append(int(GPP_title_seg_split_item[0][6:].split('_')[0]))
                        GPP_title_seg_num_dict[GPP_title_seg_split_item[0][6:].split('_')[0]] = GPP_title_seg_item

            GPP_title_seg_num_list = sorted(GPP_title_seg_num_list)
            GPP_title_half_list = [GPP_title_seg_num_dict[str(num_item)] for num_item in GPP_title_seg_num_list]
            GPP_pin_name_list.append(['GROUP ' + GPP_word[-1]] + GPP_title_half_list)
        # print(GPP_pin_name_list)

        # 找出对应的pin location
        GPP_pin_location_list = []
        pin_name_pin_location_dict = extractIOData.extract_pstchip(func='get_pin_name_pin_location')
        # print(pin_name_pin_location_dict)

        # 生成pin location list
        for GPP_pin_name_item_list in GPP_pin_name_list:
            pin_location_seg_list = [pin_name_pin_location_dict[GPP_pin_name_item] for
                                     GPP_pin_name_item in GPP_pin_name_item_list[1:]]
            GPP_pin_location_list.append([''] + pin_location_seg_list)

        # 展平list
        GPP_pin_name_list = flatten(GPP_pin_name_list)
        GPP_pin_location_list = flatten(GPP_pin_location_list)
        # ***************************************************************************************************

    # *****************生成 signal name, resistance, pu/pd, power rail***********************************
    # # 但是输入的pin脚名称是不确定的，需要进行一些处理，
    # pin_name_list = extractIOData.extract_pstchip(GPIO_pin_name_list_org=GPIO_pin_name_list_org,
    #                                               func='reshape_pin_name')
    all_net_list, net_node_list, net_node_dict = extractIOData.extract_pstxnet()
    # print(net_node_list)
    all_node_list, node_page_dict, all_res_dict, all_res_list, ic_ext_icname_dict = extractIOData.extract_pstxprt()
    extractPinData = ExtractPinData(None)
    # 输出每个IC的pin脚数目
    ext_icname_pin_num_dict = extractIOData.extract_pstchip(func='get_ic_pin_number')

    # 寻找.exp文件
    file_name = ''
    for x in os.listdir(input_path):
        (shotname, extension) = os.path.splitext(x)
        # print(extension)
        if extension == '.EXP':
            file_name = x
            file_base_name = os.path.splitext(x)[0]

    # 如果没有.exp文件，抛出异常
    if file_name == '':
        error_message = 'Missing *.EXP file'
        create_error_message(output_excel_path, error_message)
        raise FileNotFoundError

    # 生成ic BOM列表
    ic_ni_dict = extractIOData.extract_exp(file_name)
    # print(ic_ni_dict)
    # 获得所有电源线
    Exclude_Net_List, PWR_Net_List, GND_Net_List = get_exclude_netlist(all_net_list)
    IC_pin_num_dict = {}
    for ic_item in ic_ext_icname_dict.keys():
        IC_pin_num_dict[ic_item] = ext_icname_pin_num_dict[ic_ext_icname_dict.get(ic_item)]
    # 获取详细的走线信息
    pin_net_node_list, pin_net_node_dict, error_all_list = extractPinData.get_detail_layout_info(
        net_node_list, GPP_pin_name_list, all_res_list, IC_pin_num_dict, Exclude_Net_List, ic_ni_dict)
    # print(all_res_list)
    # 对含有NI的数据进行删减， NI表示不上电
    for key_item in pin_net_node_dict.keys():
        for value_item in pin_net_node_dict[key_item]:
            if 'NI' in value_item:
                pin_net_node_dict[key_item].remove(value_item)

    # 对pin_net_node_dict中的value进行去重
    for key_item in pin_net_node_dict.keys():
        pin_net_node_dict[key_item] = list(set([tuple(t) for t in pin_net_node_dict[key_item]]))

    # 通过GPI/O这一项的信息来筛选详细走线信息
    real_signal_name_list = []
    real_resistance_list = []
    real_pu_pd_list = []
    real_power_list = []

    # ************************************************ 生成result及错误信息 ***************************************
    # 生成错误信息
    result_list = []
    error_message_list = []
    # print(all_res_dict)
    # print(1, len(GPP_pin_name_list))
    for gpp_idx in range(len(GPP_pin_name_list)):
        pin_name_item = GPP_pin_name_list[gpp_idx]
        pwr_num = 0
        gnd_num = 0
        pwr_item = ''
        gnd_item = ''
        pwr_rail_list = []
        gnd_rail_list = []
        pwr_item_list = []
        gnd_item_list = []
        pwr_res_item_list = []
        gnd_res_item_list = []
        # print(pin_name_item)
        if pin_name_item.upper().find('GROUP') == -1:
            net_node_item = pin_net_node_dict[pin_name_item]
            try:
                # 判断是上拉还是下拉还是分压
                for net_item in net_node_item:
                    # 如果找到了上拉电源
                    if net_item[-1] in PWR_Net_List:
                        pwr_num += 1
                        rail_item = ', ' + net_item[-1]
                        pwr_item += rail_item if pwr_num > 1 else net_item[-1]
                        pwr_rail_list.append(net_item[-1])
                        pwr_item_list.append(net_item[:-2])
                        if net_item[-2] in all_res_list:
                            pwr_res_item_list.append(all_res_dict[net_item[-2]])
                        elif net_item[-4] in all_res_list:
                            pwr_res_item_list.append(all_res_dict[net_item[-3]])
                        else:
                            pwr_res_item_list.append(None)

                    # 找到了下拉地线
                    if net_item[-1] in GND_Net_List:
                        rail_item = ', ' + net_item[-1]
                        gnd_num += 1
                        gnd_item += rail_item if pwr_num > 1 else net_item[-1]
                        gnd_rail_list.append(net_item[-1])
                        gnd_item_list.append(net_item[:-2])
                        if net_item[-2] in all_res_list:
                            gnd_res_item_list.append(all_res_dict[net_item[-2]])
                        elif net_item[-4] in all_res_list:
                            gnd_res_item_list.append(all_res_dict[net_item[-4]])
                        else:
                            gnd_res_item_list.append(None)

                real_signal_name_list.append(net_node_item[0][0])

                # 出现三条以上电源线，则报错
                if pwr_num > 2:
                    real_resistance_list.append(None)
                    real_pu_pd_list.append(None)
                    real_power_list.append(pwr_item)
                    result_list.append('Fail')
                    error_message_list.append("The number of power rails exceeds three")
                # 出现三条以上地线，则报错
                elif gnd_num > 2:
                    real_resistance_list.append(None)
                    real_pu_pd_list.append(None)
                    real_power_list.append(gnd_item)
                    result_list.append('Fail')
                    error_message_list.append("The number of ground lines exceeds three")
                # 如果既有电源线又有地线，则判断为分压
                elif pwr_num > 0 and gnd_num > 0:
                    if gnd_res_item_list[0] not in [None, '0'] and pwr_res_item_list[0] not in [None, '0']:
                        if pwr_item_list[0] == gnd_item_list[0]:
                            real_resistance_list.append(pwr_res_item_list[0])
                            real_pu_pd_list.append('PU/PD')
                            real_power_list.append(pwr_item)
                            result_list.append('Check')
                            error_message_list.append(None)
                        # 不一样就是上拉
                        else:
                            real_resistance_list.append(pwr_res_item_list[0] + ', ' + gnd_res_item_list[0])
                            real_pu_pd_list.append('PU/PD')
                            real_power_list.append(pwr_item + ', ' + gnd_item)
                            result_list.append('Fail')
                            error_message_list.append('There are both power rail and ground rail')

                    # 如果为None说明没有电阻
                    else:
                        real_resistance_list.append(pwr_res_item_list[0] + ', ' + gnd_res_item_list[0])
                        real_pu_pd_list.append('PU/PD')
                        real_power_list.append(pwr_item + ', ' + gnd_item)
                        result_list.append('Fail')
                        error_message_list.append('There are both power rail and ground rail')

                # 出现两条电源线，先判断是不是接的同名电源线
                elif pwr_num > 1:
                    # 如果有两个电阻
                    if pwr_res_item_list[0] not in [None, '0'] and pwr_res_item_list[1] not in [None, '0']:
                        # 如果两个电阻相同, 说明是同一根电源线
                        if pwr_res_item_list[0] == pwr_res_item_list[-1]:
                            real_resistance_list.append(pwr_res_item_list[0])
                            real_pu_pd_list.append('PU')
                            real_power_list.append(pwr_rail_list[0])
                            result_list.append('Pass')
                            error_message_list.append(None)
                        # 如果两个电阻不同，说明是不同电源线
                        elif pwr_rail_list[0] == pwr_rail_list[1]:
                            real_resistance_list.append(pwr_res_item_list[0] + ', ' + pwr_res_item_list[1])
                            real_pu_pd_list.append('PU')
                            real_power_list.append(pwr_rail_list[0])
                            result_list.append('Fail')
                            error_message_list.append("There are two different resistors")
                        else:
                            real_resistance_list.append(pwr_res_item_list[0] + ', ' + pwr_res_item_list[1])
                            real_pu_pd_list.append('PU')
                            real_power_list.append(pwr_item)
                            result_list.append('Fail')
                            error_message_list.append("There are two different power rails")
                    # 只有一个电阻
                    else:
                        if pwr_res_item_list[0] not in [None, '0']:
                            real_resistance_list.append(pwr_res_item_list[0])
                        else:
                            real_resistance_list.append(pwr_res_item_list[1])
                        real_pu_pd_list.append('PU')
                        real_power_list.append(pwr_item)
                        result_list.append('Fail')
                        error_message_list.append("There are two different power rails")

                # 出现两条地线
                elif gnd_num > 1:
                    # 如果有两个电阻
                    if gnd_res_item_list[0] not in [None, '0'] and gnd_res_item_list[1] not in [None, '0']:
                        # 如果两个电阻相同, 说明是同一根地线
                        if gnd_res_item_list[0] == gnd_res_item_list[-1]:
                            real_resistance_list.append(gnd_res_item_list[0])
                            real_pu_pd_list.append('PD')
                            real_power_list.append(gnd_rail_list[0])
                            result_list.append('Pass')
                            error_message_list.append(None)
                        # 如果两个电阻不同，说明是不同地线
                        else:
                            real_resistance_list.append(gnd_res_item_list[0] + ', ' + gnd_res_item_list[1])
                            real_pu_pd_list.append('PD')
                            real_power_list.append(gnd_item)
                            result_list.append('Fail')
                            error_message_list.append("There are two different resistors")
                    # 如果只有一个电阻
                    else:
                        if gnd_res_item_list[0] not in [None, '0']:
                            real_resistance_list.append(gnd_res_item_list[0])
                        else:
                            real_resistance_list.append(gnd_res_item_list[1])
                        real_pu_pd_list.append('PD')
                        real_power_list.append(gnd_item)
                        result_list.append('Fail')
                        error_message_list.append("The number of ground lines exceeds two")
                # 上拉PU
                elif pwr_num == 1:
                    if pwr_res_item_list[0] == '0':
                        real_resistance_list.append(None)
                        real_pu_pd_list.append(None)
                        real_power_list.append(None)
                        result_list.append('Pass')
                        error_message_list.append(None)
                    else:
                        if pwr_res_item_list[0] is not None:
                            real_resistance_list.append(pwr_res_item_list[0])
                            real_pu_pd_list.append('PU')
                            real_power_list.append(pwr_item)
                            result_list.append('Pass')
                            error_message_list.append(None)
                        else:
                            real_resistance_list.append(None)
                            real_pu_pd_list.append('PU')
                            real_power_list.append(pwr_item)
                            result_list.append('Fail')
                            error_message_list.append('Lack of pull_up resistor')
                # 下拉PD
                elif gnd_num == 1:
                    if gnd_res_item_list[0] == '0':
                        real_resistance_list.append(None)
                        real_pu_pd_list.append(None)
                        real_power_list.append(None)
                        result_list.append('Pass')
                        error_message_list.append(None)
                    else:
                        if gnd_res_item_list[0] is not None:
                            real_resistance_list.append(gnd_res_item_list[0])
                            real_pu_pd_list.append('PD')
                            real_power_list.append(gnd_item)
                            result_list.append('Pass')
                            error_message_list.append(None)
                        else:
                            real_resistance_list.append(None)
                            real_pu_pd_list.append('PD')
                            real_power_list.append(gnd_item)
                            result_list.append('Fail')
                            error_message_list.append('Lack of pull_down resistor')
                # 既没有上拉也没有下拉
                else:
                    real_resistance_list.append(None)
                    real_pu_pd_list.append(None)
                    real_power_list.append(None)
                    result_list.append('Pass')
                    error_message_list.append(None)

            except IndexError:
                # print(None)
                real_resistance_list.append(None)
                real_pu_pd_list.append(None)
                real_power_list.append(None)
                result_list.append('Pass')
                error_message_list.append(None)
                real_signal_name_list.append('NC')
        else:
            # print(pin_name_item)
            # 如果是GROUP
            real_signal_name_list.append(None)
            real_resistance_list.append(None)
            real_pu_pd_list.append(None)
            real_power_list.append(None)
            result_list.append(None)
            error_message_list.append(None)

    # 如果是更新数据
    if excel_exist_flag:
        pre_content_list = [x[-6:-3] for x in content[2:]]

        compare_resistance_list = real_resistance_list[1:]
        compare_pu_pd_list = real_pu_pd_list[1:]
        compare_power_list = real_power_list[1:]

        now_content_list = [[compare_resistance_list[x_idx], compare_pu_pd_list[x_idx],
                             compare_power_list[x_idx]] for x_idx in range(len(compare_resistance_list))]

        compare_result_list = []
        compare_error_message_list = []

        for com_idx in range(len(pre_content_list)):

            pre_item = pre_content_list[com_idx]
            now_item = now_content_list[com_idx]
            resistance_flag = False
            pu_pd_flag = False
            power_flag = False

            # resistance不同
            if pre_item[0] != now_item[0]:
                resistance_flag = True
            # pu/pd不同
            if pre_item[1] != now_item[1]:
                pu_pd_flag = True
            # power不同
            if pre_item[2] != now_item[2]:
                power_flag = True

            if resistance_flag and pu_pd_flag and power_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('All items are inconsistent with before')
            elif resistance_flag and pu_pd_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('Resistance and PU/PD are inconsistent with before')
            elif resistance_flag and power_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('Resistance and Power are inconsistent with before')
            elif pu_pd_flag and power_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('PU/PD and Power are inconsistent with before')
            elif resistance_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('Resistance is inconsistent with before')
            elif pu_pd_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('PU/PD is inconsistent with before')
            elif power_flag:
                compare_result_list.append('Fail')
                compare_error_message_list.append('Power is inconsistent with before')
            else:
                compare_result_list.append('Pass')
                compare_error_message_list.append(None)

    # *************************************************************************************************
    # GPP_pin_name_list = np.array(GPP_pin_name_list).reshape(-1, 1).tolist()
    # GPP_pin_location_list = np.array(GPP_pin_location_list).reshape(-1, 1).tolist()
    # ******************************创建表格并写入******************************************************
    # 创建ouput文件夹和output表格
    if os.path.exists(output_path):
        pass
    else:
        os.mkdir(output_path)

    # 列项信息列表
    column_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
                   'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                   'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ',
                   'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                   'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ',
                   'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
                   'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ',
                   'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ',
                   'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP', 'DQ',
                   'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ',
                   'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP', 'EQ',
                   'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ',
                   'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP', 'FQ',
                   'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ',
                   'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP', 'GQ',
                   'GR', 'GS', 'GT', 'GU', 'GV', 'GW', 'GX', 'GY', 'GZ',
                   'HA', 'HB', 'HC', 'HD', 'HE', 'HF', 'HG', 'HH', 'HI', 'HJ', 'HK', 'HL', 'HM', 'HN', 'HO', 'HP', 'HQ',
                   'HR', 'HS', 'HT', 'HU', 'HV', 'HW', 'HX', 'HY', 'HZ',
                   'IA', 'IB', 'IC', 'ID', 'IE', 'IF', 'IG', 'IH', 'II', 'IJ', 'IK', 'IL', 'IM', 'IN', 'IO', 'IP', 'IQ',
                   'IR', 'IS', 'IT', 'IU', 'IV', 'IW', 'IX', 'IY', 'IZ',
                   'JA', 'JB', 'JC', 'JD', 'JE', 'JF', 'JG', 'JH', 'JI', 'JJ', 'JK', 'JL', 'JM', 'JN', 'JO', 'JP', 'JQ',
                   'JR', 'JS', 'JT', 'JU', 'JV', 'JW', 'JX', 'JY', 'JZ',
                   'KA', 'KB', 'KC', 'KD', 'KE', 'KF', 'KG', 'KH', 'KI', 'KJ', 'KK', 'KL', 'KM', 'KN', 'KO', 'KP', 'KQ',
                   'KR', 'KS', 'KT', 'KU', 'KV', 'KW', 'KX', 'KY', 'KZ',
                   'LA', 'LB', 'LC', 'LD', 'LE', 'LF', 'LG', 'LH', 'LI', 'LJ', 'LK', 'LL', 'LM', 'LN', 'LO', 'LP', 'LQ',
                   'LR', 'LS', 'LT', 'LU', 'LV', 'LW', 'LX', 'LY', 'LZ',
                   'MA', 'MB', 'MC', 'MD', 'ME', 'MF', 'MG', 'MH', 'MI', 'MJ', 'MK', 'ML', 'MM', 'MN', 'MO', 'MP', 'MQ',
                   'MR', 'MS', 'MT', 'MU', 'MV', 'MW', 'MX', 'MY', 'MZ',
                   'NA', 'NB', 'NC', 'ND', 'NE', 'NF', 'NG', 'NH', 'NI', 'NJ', 'NK', 'NL', 'NM', 'NN', 'NO', 'NP', 'NQ',
                   'NR', 'NS', 'NT', 'NU', 'NV', 'NW', 'NX', 'NY', 'NZ',
                   'OA', 'OB', 'OC', 'OD', 'OE', 'OF', 'OG', 'OH', 'OI', 'OJ', 'OK', 'OL', 'OM', 'ON', 'OO', 'OP', 'OQ',
                   'OR', 'OS', 'OT', 'OU', 'OV', 'OW', 'OX', 'OY', 'OZ',
                   'PA', 'PB', 'PC', 'PD', 'PE', 'PF', 'PG', 'PH', 'PI', 'PJ', 'PK', 'PL', 'PM', 'PN', 'PO', 'PP', 'PQ',
                   'PR', 'PS', 'PT', 'PU', 'PV', 'PW', 'PX', 'PY', 'PZ',
                   'QA', 'QB', 'QC', 'QD', 'QE', 'QF', 'QG', 'QH', 'QI', 'QJ', 'QK', 'QL', 'QM', 'QN', 'QO', 'QP', 'QQ',
                   'QR', 'QS', 'QT', 'QU', 'QV', 'QW', 'QX', 'QY', 'QZ',
                   'RA', 'RB', 'RC', 'RD', 'RE', 'RF', 'RG', 'RH', 'RI', 'RJ', 'RK', 'RL', 'RM', 'RN', 'RO', 'RP', 'RQ',
                   'RR', 'RS', 'RT', 'RU', 'RV', 'RW', 'RX', 'RY', 'RZ',
                   'SA', 'SB', 'SC', 'SD', 'SE', 'SF', 'SG', 'SH', 'SI', 'SJ', 'SK', 'SL', 'SM', 'SN', 'SO', 'SP', 'SQ',
                   'SR', 'SS', 'ST', 'SU', 'SV', 'SW', 'SX', 'SY', 'SZ',
                   'TA', 'TB', 'TC', 'TD', 'TE', 'TF', 'TG', 'TH', 'TI', 'TJ', 'TK', 'TL', 'TM', 'TN', 'TO', 'TP', 'TQ',
                   'TR', 'TS', 'TT', 'TU', 'TV', 'TW', 'TX', 'TY', 'TZ',
                   'UA', 'UB', 'UC', 'UD', 'UE', 'UF', 'UG', 'UH', 'UI', 'UJ', 'UK', 'UL', 'UM', 'UN', 'UO', 'UP', 'UQ',
                   'UR', 'US', 'UT', 'UU', 'UV', 'UW', 'UX', 'UY', 'UZ',
                   'VA', 'VB', 'VC', 'VD', 'VE', 'VF', 'VG', 'VH', 'VI', 'VJ', 'VK', 'VL', 'VM', 'VN', 'VO', 'VP', 'VQ',
                   'VR', 'VS', 'VT', 'VU', 'VV', 'VW', 'VX', 'VY', 'VZ',
                   'WA', 'WB', 'WC', 'WD', 'WE', 'WF', 'WG', 'WH', 'WI', 'WJ', 'WK', 'WL', 'WM', 'WN', 'WO', 'WP', 'WQ',
                   'WR', 'WS', 'WT', 'WU', 'WV', 'WW', 'WX', 'WY', 'WZ',
                   'XA', 'XB', 'XC', 'XD', 'XE', 'XF', 'XG', 'XH', 'XI', 'XJ', 'XK', 'XL', 'XM', 'XN', 'XO', 'XP', 'XQ',
                   'XR', 'XS', 'XT', 'XU', 'XV', 'XW', 'XX', 'XY', 'XZ',
                   'YA', 'YB', 'YC', 'YD', 'YE', 'YF', 'YG', 'YH', 'YI', 'YJ', 'YK', 'YL', 'YM', 'YN', 'YO', 'YP', 'YQ',
                   'YR', 'YS', 'YT', 'YU', 'YV', 'YW', 'YX', 'YY', 'YZ',
                   'ZA', 'ZB', 'ZC', 'ZD', 'ZE', 'ZF', 'ZG', 'ZH', 'ZI', 'ZJ', 'ZK', 'ZL', 'ZM', 'ZN', 'ZO', 'ZP', 'ZQ',
                   'ZR', 'ZS', 'ZT', 'ZU', 'ZV', 'ZW', 'ZX', 'ZY', 'ZZ'
                   ]

    # 如果存在说明是更新版本，直接向后添加
    if excel_exist_flag:

        title_list = ['GPI/O', 'GPIO Result', 'Signal name', 'Resistance', 'PU/PD', 'Power Rail', 'Result',
                      'Error Message', 'Remark']

        GPIO_result_list = []
        result_color_list = []
        # 生成对应的result_list公式
        for idx in range(len(compare_result_list)):
            if compare_result_list[idx]:
                if compare_result_list[idx] == 'Pass':
                    result_color_list.append(0)
                elif compare_result_list[idx] == 'Check':
                    result_color_list.append(1)
                else:
                    result_color_list.append(2)
                GPIO_result_list.append('=IF(OR(AND({}{}="GPI",{}{}="PU"),AND({}{}="GPO",{}{}="PD"),'
                                        '{}{}="",{}{}="NEGATIVE"),"Pass","Fail")'
                                        .format(column_list[column_length], idx + 4, column_list[column_length + 4],
                                                idx + 4, column_list[column_length], idx + 4,
                                                column_list[column_length + 4], idx + 4, column_list[column_length],
                                                idx + 4, column_list[column_length + 4], idx + 4))
            else:
                GPIO_result_list.append(None)
                result_color_list.append(None)

        # 创建新的表格
        workbook = xlsxwriter.Workbook(output_excel_path)
        worksheet = workbook.add_worksheet('GPIO_TABLE')

        # 样式设置
        title_format = workbook.add_format({'font_size': 16, 'bold': True, 'bg_color': '#9BC2E6', 'border': 1})
        column_pin_format = workbook.add_format({'font_size': 16, 'bg_color': '#65DADD', 'border': 1})
        column_gpio_format = workbook.add_format({'font_size': 16, 'bg_color': '#DAEEF3', 'border': 1, 'locked': 0})
        column_pre_gpio_format = workbook.add_format({'font_size': 16, 'bg_color': '#DAEEF3', 'border': 1})
        column_content_format = workbook.add_format({'font_size': 16, 'bg_color': '#A4ECF6', 'border': 1})
        column_group_format = workbook.add_format({'font_size': 16, 'bg_color': '#FFE699', 'border': 1})
        column_pass_format = workbook.add_format({'font_size': 16, 'bg_color': '#8AF371', 'border': 1})
        column_check_format = workbook.add_format({'font_size': 16, 'bg_color': '#FFFF00', 'border': 1})
        column_fail_format = workbook.add_format({'font_size': 16, 'bg_color': '#FC2443', 'border': 1})

        count_idx = int((column_length - 2) / 9)

        # 对 GPI/O Result 设置 conditional_format
        for x_idx1 in range(len(GPIO_result_list)):
            worksheet.conditional_format('{}{}'.format(column_list[count_idx * 9 + 3], x_idx1 + 4),
                                         {'type': 'text',
                                          'criteria': 'containing',
                                          'value': 'Pass',
                                          'format': column_pass_format})
            worksheet.conditional_format('{}{}'.format(column_list[count_idx * 9 + 3], x_idx1 + 4),
                                         {'type': 'text',
                                          'criteria': 'containing',
                                          'value': 'Fail',
                                          'format': column_fail_format})

        merge_format1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 16, 'border': 1})

        merge_format2 = workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                             'bg_color': '#F99FCC', 'font_size': 22, 'bold': True, 'border': 1})

        # 生成column_list
        real_gpio_list = ['GPI/O'] + [None] * (len(GPP_pin_name_list) - 1)
        real_gpio_result_list = ['GPI/O Result', None] + [None] * (len(GPP_pin_name_list) - 1)
        real_signal_name_list = ['Signal name'] + real_signal_name_list
        real_resistance_list = ['Resistance'] + real_resistance_list
        real_pu_pd_list = ['PU/PD'] + real_pu_pd_list
        real_power_list = ['Power Rail'] + real_power_list
        real_result_list = ['Result', None] + compare_result_list
        error_message_list = ['Error Message', None] + compare_error_message_list
        real_remark_list = ['Remark'] + [None] * (len(GPP_pin_name_list) - 1)

        columns_list = columns_list + [real_gpio_list, real_gpio_result_list, real_signal_name_list,
                                       real_resistance_list, real_pu_pd_list, real_power_list, real_result_list,
                                       error_message_list, real_remark_list]

        # 对更新前result数据进行处理
        all_result_list = [columns_list[x * 9 + 8][2:] for x in range(count_idx)]
        all_gpio_result_list = [columns_list[x * 9 + 3][2:] for x in range(count_idx)]
        all_result_color_list = []
        all_gpio_result_color_list = []

        for all_item in all_gpio_result_list:
            gpio_color_list = []
            for item in all_item:
                if item:
                    if item == 'Pass':
                        gpio_color_list.append(0)
                    else:
                        gpio_color_list.append(1)
                else:
                    gpio_color_list.append(None)

            all_gpio_result_color_list.append(gpio_color_list)

        for all_item in all_result_list:
            color_list = []
            for item1 in all_item:
                if item1:
                    if item1 == 'Pass':
                        color_list.append(0)
                    elif item1 == 'Check':
                        color_list.append(1)
                    else:
                        color_list.append(2)
                else:
                    color_list.append(None)

            all_result_color_list.append(color_list)
        # ********************************************* 上色 **************************************************
        # 对数据进行写入上色
        for col_idx in range(len(columns_list)):
            # print(col_idx + 1, columns_list[col_idx])
            # 前两个上pin的颜色
            if col_idx in [0, 1]:
                worksheet.write_column('{}2'.format(column_list[col_idx]), columns_list[col_idx], column_pin_format)
            # 对之前的GPIO,REMARK上色
            elif col_idx in [x * 9 + 2 for x in range(count_idx)] + [x * 9 + 10 for x in range(count_idx)]:
                worksheet.write_column('{}2'.format(column_list[col_idx]), columns_list[col_idx],
                                       column_pre_gpio_format)
            # 其他上content颜色
            else:
                worksheet.write_column('{}2'.format(column_list[col_idx]), columns_list[col_idx], column_content_format)

        # 对更新的gpio result上色
        worksheet.write_column('{}4'.format(column_list[column_length + 1]), GPIO_result_list, column_pass_format)

        # 对result进行上色
        for gpio_idx1 in range(len(all_gpio_result_color_list)):
            gpio_item = all_gpio_result_color_list[gpio_idx1]
            for gpio_idx2 in range(len(gpio_item)):
                item = gpio_item[gpio_idx2]
                if item == 0:
                    worksheet.write('{}{}'.format(column_list[gpio_idx1 * 9 + 3], 4 + gpio_idx2),
                                    all_gpio_result_list[gpio_idx1][gpio_idx2], column_pass_format)
                # fail
                elif item == 1:
                    worksheet.write('{}{}'.format(column_list[gpio_idx1 * 9 + 3], 4 + gpio_idx2),
                                    all_gpio_result_list[gpio_idx1][gpio_idx2], column_fail_format)
                else:
                    pass

        for result_idx1 in range(len(all_result_color_list)):
            result_item = all_result_color_list[result_idx1]
            for result_idx2 in range(len(result_item)):
                item = result_item[result_idx2]
                # pass
                if item == 0:
                    worksheet.write('{}{}'.format(column_list[result_idx1 * 9 + 8], 4 + result_idx2),
                                    all_result_list[result_idx1][result_idx2], column_pass_format)
                # check
                elif item == 1:
                    worksheet.write('{}{}'.format(column_list[result_idx1 * 9 + 8], 4 + result_idx2),
                                    all_result_list[result_idx1][result_idx2], column_check_format)
                elif item == 2:
                    worksheet.write('{}{}'.format(column_list[result_idx1 * 9 + 8], 4 + result_idx2),
                                    all_result_list[result_idx1][result_idx2], column_fail_format)
                else:
                    pass

        # 循环生成并设置版本信息及颜色
        # 设置标题信息及颜色
        worksheet.merge_range('A1:B1', file_base_name, merge_format1)
        for con_item in range(count_idx + 1):
            worksheet.merge_range('{}1:{}1'.format(column_list[con_item * 9 + 2], column_list[con_item * 9 + 9]),
                                  'X' + str("%02d" % con_item), merge_format2)

        # 对更新的result上色
        for result_idx in range(len(result_color_list)):
            if result_color_list[result_idx] == 0:
                worksheet.write('{}{}'.format(column_list[column_length + 6], result_idx + 4),
                                compare_result_list[result_idx], column_pass_format)
            # check
            elif result_color_list[result_idx] == 1:
                worksheet.write('{}{}'.format(column_list[column_length + 6], result_idx + 4),
                                compare_result_list[result_idx], column_check_format)
            # fail
            elif result_color_list[result_idx] == 2:
                worksheet.write('{}{}'.format(column_list[column_length + 6], result_idx + 4),
                                compare_result_list[result_idx], column_fail_format)
            else:
                pass

        # 给title上色
        all_title_list = content[0] + title_list
        worksheet.write_row('A2', all_title_list, title_format)

        # 对表格进行自适应
        set_column_width(columns_list, worksheet)

        # 设置工作表保护
        worksheet.protect('Gorgeous')

        # 设置不保护的项: GPI/0 和 Remark
        worksheet.write_column('{}4'.format(column_list[column_length]),
                               [None] * (len(GPP_pin_name_list) - 1), column_gpio_format)
        worksheet.write_column('{}4'.format(column_list[column_length + 8]),
                               [None] * (len(GPP_pin_name_list) - 1), column_gpio_format)

        # 改变Group的颜色
        group_idx_list = [x_idx + 3 for x_idx in range(len(GPP_pin_name_list)) if
                          GPP_pin_name_list[x_idx].find('GROUP') > -1]
        for idx in range(len(group_idx_list)):
            group_idx = group_idx_list[idx]
            group_length = len(all_title_list) - 1
            group_list = ['GROUP {}'.format(column_list[idx])] + [None] * group_length

            worksheet.write_row('A{}'.format(group_idx), group_list, column_group_format)

        # *****************************************************************************************************

        # 设置冻结窗口
        worksheet.freeze_panes(2, 2)

        workbook.close()

    # 不存在说明是第一次运行，则直接创建新的表格
    else:

        title_list = ['Pin Name', 'Pin Location', 'GPI/O', 'GPI/O Result', 'Signal name', 'Resistance', 'PU/PD',
                      'Power Rail', 'Result', 'Error Message', 'Remark']

        GPIO_result_list = []
        result_color_list = []

        # 生成对应的result_list公式
        for idx in range(len(result_list)):
            if result_list[idx]:
                if result_list[idx] == 'Pass':
                    result_color_list.append(0)
                elif result_list[idx] == 'Check':
                    result_color_list.append(1)
                else:
                    result_color_list.append(2)
                GPIO_result_list.append('=IF(OR(AND(C{}="GPI",G{}="PU"),AND(C{}="GPO",G{}="PD"), C{}="",'
                                        'C{}="NEGATIVE"),"Pass","Fail")'
                                        .format(idx + 3, idx + 3, idx + 3, idx + 3, idx + 3, idx + 3))
            else:
                GPIO_result_list.append(None)
                result_color_list.append(None)

        # 創建excel
        workbook = xlsxwriter.Workbook(output_excel_path)
        worksheet = workbook.add_worksheet('GPIO_TABLE')

        # 合并单元格
        merge_format1 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 16, 'border': 1})
        worksheet.merge_range('A1:B1', file_base_name, merge_format1)

        merge_format2 = workbook.add_format({'align': 'center', 'valign': 'vcenter',
                                             'bg_color': '#F99FCC', 'font_size': 22, 'bold': True, 'border': 1})
        worksheet.merge_range('C1:I1', 'X00', merge_format2)
        worksheet.write('J1', None, merge_format2)
        # 写入表格数据
        title_format = workbook.add_format({'font_size': 16, 'bold': True, 'bg_color': '#9BC2E6', 'border': 1})
        column_pin_format = workbook.add_format({'font_size': 16, 'bg_color': '#65DADD', 'border': 1})
        column_gpio_format = workbook.add_format({'font_size': 16, 'bg_color': '#DAEEF3', 'border': 1, 'locked': 0})
        column_content_format = workbook.add_format({'font_size': 16, 'bg_color': '#A4ECF6', 'border': 1})
        column_group_format = workbook.add_format({'font_size': 16, 'bg_color': '#FFE699', 'border': 1})
        column_pass_format = workbook.add_format({'font_size': 16, 'bg_color': '#8AF371', 'border': 1})
        column_check_format = workbook.add_format({'font_size': 16, 'bg_color': '#FFFF00', 'border': 1})
        column_fail_format = workbook.add_format({'font_size': 16, 'bg_color': '#FC2443', 'border': 1})

        worksheet.write_row('A2', title_list, title_format)
        worksheet.write_column('A3', GPP_pin_name_list, column_pin_format)
        worksheet.write_column('B3', GPP_pin_location_list, column_pin_format)
        worksheet.write_column('D3', GPIO_result_list, column_pass_format)
        worksheet.write_column('E3', real_signal_name_list, column_content_format)
        worksheet.write_column('F3', real_resistance_list, column_content_format)
        worksheet.write_column('G3', real_pu_pd_list, column_content_format)
        worksheet.write_column('H3', real_power_list, column_content_format)
        worksheet.write_column('J3', error_message_list, column_content_format)
        # print(result_list)
        for result_idx in range(len(result_color_list)):
            if result_color_list[result_idx] == 0:
                worksheet.write('I{}'.format(result_idx + 3), result_list[result_idx], column_pass_format)
            # check
            elif result_color_list[result_idx] == 1:
                worksheet.write('I{}'.format(result_idx + 3), result_list[result_idx], column_check_format)
            # fail
            elif result_color_list[result_idx] == 2:
                worksheet.write('I{}'.format(result_idx + 3), result_list[result_idx], column_fail_format)
            else:
                pass

        # 对 GPI/O Result 设置 conditional_format
        for x_idx1 in range(len(GPIO_result_list)):
            worksheet.conditional_format('D{}'.format(x_idx1 + 4),
                                         {'type': 'text',
                                          'criteria': 'containing',
                                          'value': 'Pass',
                                          'format': column_pass_format})
            worksheet.conditional_format('D{}'.format(x_idx1 + 4),
                                         {'type': 'text',
                                          'criteria': 'containing',
                                          'value': 'Fail',
                                          'format': column_fail_format})

        # 生成column_list
        GPP_pin_name_list = ['Pin Name'] + GPP_pin_name_list
        GPP_pin_location_list = ['Pin Location'] + GPP_pin_location_list
        real_signal_name_list = ['Signal name'] + real_signal_name_list
        real_resistance_list = ['Resistance'] + real_resistance_list
        real_pu_pd_list = ['PU/PD'] + real_pu_pd_list
        real_power_list = ['Power Rail'] + real_power_list
        error_message_list = ['Error Message'] + error_message_list

        columns_list = [GPP_pin_name_list, GPP_pin_location_list, ['GPI/O'], ['GPI/O Result'], real_signal_name_list,
                        real_resistance_list, real_pu_pd_list, real_power_list, ['Result'],
                        error_message_list, ['Remark']]
        # 对表格进行自适应
        set_column_width(columns_list, worksheet)

        # 设置冻结窗口
        worksheet.freeze_panes(2, 2)

        # 设置工作表保护
        worksheet.protect('Gorgeous')

        # 设置不保护的项: GPI/0 和 Remark
        worksheet.write_column('C4', [None for _ in range(len(GPP_pin_name_list) - 2)], column_gpio_format)
        worksheet.write_column('K4', [None for _ in range(len(GPP_pin_name_list) - 2)], column_gpio_format)

        # 改变Group的颜色
        group_idx_list = [x_idx + 2 for x_idx in range(len(GPP_pin_name_list)) if
                          GPP_pin_name_list[x_idx].find('GROUP') > -1]
        for idx in range(len(group_idx_list)):
            group_idx = group_idx_list[idx]
            group_length = len(title_list) - 1
            group_list = ['GROUP {}'.format(column_list[idx])] + [None] * group_length

            worksheet.write_row('A{}'.format(group_idx), group_list, column_group_format)

        workbook.close()


# 自适应功能
def set_column_width(columns, worksheet):
    length_list = [ceil(max([len(str(y)) for y in x]) * 1.7) for x in columns]
    for i, width in enumerate(length_list):
        # print(i, i, width + 5)
        worksheet.set_column(i, i, width)


if __name__ == '__main__':
    generate_report()
