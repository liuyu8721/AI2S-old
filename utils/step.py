from __future__ import division
import sqlite3
import re

import regex
# from LAC import LAC
import requests
from urllib.parse import *
import pandas as pd
import hanlp

# from paddle.distributed.fleet import step

# # 装载LAC模型
# lac = LAC(mode='lac')

# 装载 HanLP 模型
HanLP = hanlp.load(
    hanlp.pretrained.mtl.CLOSE_TOK_POS_NER_SRL_DEP_SDP_CON_ERNIE_GRAM_ZH)

# def clean_blank_line(text):
#     """去除文本中空行"""
# clear_lines = ""
# for line in text:
#     if line != '\n':
#         clear_lines += line
# return clear_lines


def Q2B(uchar):
    """单个字符 全角转半角"""
    inside_code = ord(uchar)
    if inside_code == 12288:  # 全角空格直接转换
        inside_code = 32
    elif 65281 <= inside_code <= 65374:  # 全角字符（除空格）根据关系转化
        inside_code -= 65248
    return chr(inside_code)


def stringQ2B(ustring):
    """把字符串全角转半角"""
    return "".join([Q2B(uchar) for uchar in ustring])


def standardize_age(text):
    """将小于一岁的年龄标准化表示(年龄=月数/12)"""
    age_pattern = r'(?:,)(\d{1,2})个月(?:,)'
    match = re.search(age_pattern, text)
    if match:
        s = match.start()
        e = match.end()
        age = round(int(match.group(1)) / 12, 2)
        return text[:s] + str(age) + '岁' + text[e:]
    else:
        return text


def clean_data(text):
    """处理替换文本中行动动词、标点符号、同义词,删除无用标点、字符"""

    def go_to(text):
        '''修改替换文本中不以句号结束的行动动词，统一为‘到’'''
        pattern = '(?:抵达|返回|进入|在|回|赴|回到|前往|去到|飞往)(?!\。)'
        match = re.search(pattern, text)
        while match:
            s = match.start()
            e = match.end()
            text = text[:s] + '到' + text[e:]
            match = re.search(pattern, text)
        return text

    '''删除不表示时间的冒号,删除'--',删除空格'''
    text = re.sub(r'(?<=\D):|--|\s', '', text)
    '''删除序列标号 例如: 1. 3. 1、 2、'''
    text = re.sub(r'[1-9]\d*(?:\.(?:(?=\d月)|(?!\d))|\、)', '', text)
    '''替换逗号,后括号 为 空格'''
    text = re.sub('[,)]', ' ', text)
    '''删除前括号'''
    text = re.sub('\s?\(', '', text)
    '''替换箭头(->或→)为句号'''
    text = re.sub(r'->|→', '。', text)
    '''替换分号为句号'''
    text = re.sub(r';', '。', text)
    '''删除无用干扰文字'''
    text = re.sub(r'微信关注椰城', '', text)
    '''将因前面处理后生成的句号而导致两个连续句号产生,替换为一个句号'''
    text = re.sub(r'。。', '。', text)
    '''修改替换文本中不以句号结束的行动动词，统一为‘到’'''
    text = go_to(text)
    '''替换文本中表示乘坐的动词，统一为‘乘’'''
    text = re.sub(r'乘坐|搭乘|转乘|坐', '乘', text)
    '''统一文本中火车的两种表示方式(动车或列车)为‘火车’'''
    text = re.sub(r'动车|列车', '火车', text)
    '''替换文本中‘飞机’为‘航班’'''
    text = re.sub(r'飞机', '航班', text)
    return text


def set_time(text):
    '''修改替换文本中连续day的表示 例如: 2-7日 修改为 2日-7日'''
    ret = text
    pattern = '(\d{1,2})-(\d{1,2}日)'
    match = re.search(pattern, text)
    while match:
        s = match.start()
        e = match.end()
        day = match.group(1) + '日-' + match.group(2)
        ret = ret[:s] + day + ret[e:]
        match = re.search(pattern, ret)
    return ret


def patch_month(text):
    '''补全月份:将文本中没有月份表示的日期补全为带有月份的日期
    通过查找不含月份日期文本的前文或后文中含有月份表示的最近日期,比较日的大小,判断补全月份(月份不变或+1或-1)
    例如: 1月23日到琼山区龙塘镇仁三村委会道本村 25日上午11时从龙塘镇道本村到海口府城开始出车拉客 补全为:
    1月23日到琼山区龙塘镇仁三村委会道本村 1月25日上午11时从龙塘镇道本村到海口府城开始出车拉客'''
    ret = text
    pattern = '(?<!月|\d)(\d{1,2})日'
    match = re.search(pattern, text)
    while match:
        day = match.group(1)
        s = match.start()
        e = match.end()
        prev_text = ret[:s]  # 不含月份的日期文本的前文
        next_text = ret[e:]  # 不含月份的日期文本的后文
        ptime = re.findall(r'(\d{1,2})月(\d{1,2})日',
                           prev_text)  # 查找前文是否含有带有月份的日期
        ntime = re.findall(r'(\d{1,2})月(\d{1,2})日',
                           next_text)  # 查找后文是否含有带有月份的日期
        if ptime:
            '''查找到前文距离需补全日期最近的含月日期,比较日期'''
            if int(ptime[-1][1]) <= int(day):
                month = int(ptime[-1][0])  # 前文日期小于等于需补全日期，则需补全的月即为前文的月
            else:
                month = int(ptime[-1][0]) + 1  # 前文日期大于需补全日期，则需补全的月即为前文的月+1
            date = f'{month}月{day}日'
            ret = ret[:s] + date + ret[e:]
            match = re.search(pattern, ret)
        elif ntime:
            '''查找到后文距离需补全日期最近的含月日期,比较日期'''
            if int(ntime[0][1]) < int(day):
                month = int(ntime[0][0]) - 1  # 后文日期小于需补全日期，则需补全的月即为后文的月-1
            else:
                month = int(ntime[0][0])  # 后文日期大于等于需补全日期，则需补全的月即为后文的月
            date = f'{month}月{day}日'
            ret = ret[:s] + date + ret[e:]
            match = re.search(pattern, ret)
        else:
            break
    return ret


def standardize_time(text):
    """
    标准化时间段表示:
    例如：
    1月24日中午-1月26日               →      1月24日中午~1月26日
    2019年12月18日-2020年1月19日上午  →      2019年12月18日~2020年1月19日上午
    1月25日晚-1月27日                 →       1月25日晚~1月27日
    2月13日晚至2月16日                →       2月13日晚~2月16日
    1月21-23日晚                     →       1月21日~1月23日晚
    1月20至24日                      →       1月20日~1月24日
    1月28日至30日                    →       1月28日~1月30日
    1月26日-2月1日                   →       1月26日~2月1日
    2020年1月25日下午至2月1日         →       2020年1月25日下午~2月1日
    """
    time_pattern = r'(\d{4})年?(\d{1,2})月(\d{1,2})日?(上午|中午|下午|晚)?'
    """
    pattern 匹配的组如下：
    | Group # |   Description      |
    |---------|--------------------|
    | 1       | start: year        |
    | 2       | start: month       |
    | 3       | start: day         |
    | 4       | start: time of day |
    | 5       | end: year          |
    | 6       | end: month         |
    | 7       | end: day           |
    | 8       | end: time of day   |
    """
    pattern = time_pattern + r'[-至]' + time_pattern
    match = re.search(pattern, text)

    def retut(var):
        """判断变量是否为None,不为None则返回变量,否则返回空字符"""
        return var if var else ''

    while match:
        '''提取时间表示中可能含有也可能残缺的信息:'''
        s_year = retut(match.group(1))  # 开始年份
        e_year = retut(match.group(5))  # 结束年份
        s_time = retut(match.group(4))  # 开始某日具体时间(上午|中午|晚)
        e_time = retut(match.group(8))  # 结束某日具体时间(上午|中午|晚)
        e_month = match.group(6)

        '''提取时间表示中一定含有的信息：'''
        s_month = match.group(2)
        s_day = match.group(3)
        e_day = match.group(7)

        s = match.start()
        e = match.end()

        s_desc = f'{s_year}年{s_month}月{s_day}日{s_time}' if s_year else '{s_month}月{s_day}日{s_time}'
        if e_month:
            '''处理含有两个月份的时间段:例如:2月13日至2月16日'''
            # 'time = '$'+s_year+match.group(2)+'月'+match.group(3)+'日'+s_time+'~'+e_year+match.group(6)+'月'+match.group(7)+'日'+e_time+'$'
            e_desc = f'{e_year}年{e_month}月{e_day}日{e_time}' if e_year else f'{e_month}月{e_day}日{e_time}'
        else:
            '''处理只含有一个月份的时间段:例如:1月28日至30日'''
            # time = '$'+s_year+match.group(2)+'月'+match.group(3)+'日'+s_time+'~'+e_year+match.group(2)+'月'+match.group(6)+'日'+e_time+'$'
            e_desc = f'{s_year}年{s_month}月{e_day}日{e_time}' if s_year else f'{s_month}月{e_day}日{e_time}'

        time = f'{s_desc}~{e_desc}'
        text = text[:s] + time + text[e:]
        print(text)
    return text


def timech(text):
    '''
    标准化时分表示:
    例如:
    13:30  →   13时30分
    16:00  →   16时
    '''
    pattern = r'(\d{1,2}):(\d{1,2})'
    match = re.search(pattern, text)
    while match:
        if match.group(2) == '00':
            hour = match.group(1)
            # time = '$'+hour+'时'+'$'
            time = hour + '时'
        else:
            hour = match.group(1)
            minute = match.group(2)
            # time = '$'+hour+'时'+minute+'分' + '$'
            time = hour + '时' + minute + '分'
        text = re.sub(pattern, time, text, count=1)
        match = re.search(pattern, text)
    text = re.sub('分分', '分', text)
    return text


def IsExistTrainAndFlight(text):
    '''寻找列车号和航班号,通过数据库查找是否存在'''
    TrainPattern = '(?:车次|火车)\s?([GCDZTSPKXLY1-9]\d{1,4})|乘([GCDZTSPKXLY1-9]\d{1,4})次火车'
    FlightPattern = '(?:国航|南航|海航|航班|航班号)\s?([A-Z\d]{2}\d{3,4})|([A-Z\d]{2}\d{3,4})(?:次航班|航班)'

    def LoadDataBase(txts):
        '''加载数据库'''
        f = open(txts, encoding='utf-8')
        txt = []
        for line in f:
            txt.append(line.strip())
        texts = ';'.join(txt)
        return texts

    # 加载列车号数据库
    TrainNumberDataBases = LoadDataBase('./train_number.txt')
    # 加载航班号数据库
    FlightNumberDataBases = LoadDataBase('./flight.txt')

    # 寻找
    def find(type, Pattern, DataBase, text):
        '''
        type:编号类型(火车车次或航班)
        Pattern:查找模型(火车模型或航班模型)
        DataBase:加载数据库
        text:需查找的文本
        '''
        for Match in re.finditer(Pattern, text):
            s = Match.start()
            if text[s - 1:s] == ' ':  # 删除航班或火车号前面的空格
                text = text[:s - 1] + text[s:]
            if Match.group(1) is not None:
                Number = Match.group(1)
            else:
                Number = Match.group(2)
            if re.search(Number, DataBase) == None:
                print(type + Number + ' not Exist')  # 查找数据库不存在的列车号和航班号

    find('火车车次:', TrainPattern, TrainNumberDataBases, text)
    find('航班号:', FlightPattern, FlightNumberDataBases, text)
    return text


def ModifyBlurTime(text):
    '''
    模糊时间具体化:
    通过查找模糊时间词前面的日期,修改替换模糊时间词
    例如:
    2020年1月22日乘MU2527次航班从武汉到三亚→当晚乘私家车到三亚市天涯区水蛟村住处
    →
    2020年1月22日乘MU2527次航班从武汉到三亚→2020年1月15日晚上乘私家车到三亚市天涯区水蛟村住处
    '''
    Pattern = '当天|当晚'
    match = re.search(Pattern, text)
    while match:
        time = re.findall('((?:\d{4}年)?\d{1,2}月\d{1,2}日)(?=.+(?:当天|当晚))', text)
        s = match.start()
        e = match.end()
        if time:
            if match.group() == '当晚':
                # text = text[:s]+'$'+str(time[-1])+'晚上'+'$'+text[e:]
                text = text[:s] + str(time[-1]) + '晚上' + text[e:]
            else:
                text = text[:s] + str(time[-1]) + text[e:]
            # text = text[:s]+'$'+str(time[-1])+'$'+text[e:]
            match = re.search(Pattern, text)
        else:
            break
    return text


###############################STEP 1##################################
# format text
def step1(orig_text):
    """对文本进行预处理"""
    line = []
    # 去除文本中空行
    text = orig_text.replace('\n', '')
    # 字符串全角转半角
    text = stringQ2B(text)
    # 将小于一岁的年龄标准化表示
    text = standardize_age(text)
    # 处理替换文本中行动动词、标点符号、同义词,删除无用标点、字符
    text = clean_data(text)
    # 例如:2-7日 修改为 2日-7日
    text = set_time(text)
    # 补全月份
    text = patch_month(text)
    # 标准化时间段表示
    text = standardize_time(text)
    # text = ModifyBlurTime(text)
    # 标准化时分表示
    text = timech(text)
    # 寻找列车号和航班号
    text = IsExistTrainAndFlight(text)

    # line.append(orig_text)
    line.append(text)
    return line


#######################################################################

def clean_lines(Lines):
    '''去除分隔后小于等于2个字符的段落'''
    ClearLines = []
    for line in Lines:
        if len(line) > 2:
            ClearLines.append(line)
    return ClearLines


def split_by_stop(orig_text):
    '''
    以文本中句号分隔文本
    excelfile_path:需识别的excel文档地址
    Save_textpath_ByStop:分隔处理后生成的text文本地址
    '''
    res = []
    point = re.search('。', orig_text)
    while point:
        s = point.start()
        e = point.end()
        var = orig_text[:s]
        res.append(var)
        orig_text = orig_text[e:]
        point = re.search('。', orig_text)
    return res


def split_by_date(file_path):
    '''以日期分隔文本'''
    LinesList = SimpleSplitByDate(file_path)
    LinesList = MoreSplitByDate(LinesList)
    # CDLines = SplitByCTNDate(DLines)
    LinesList = GetToDate(LinesList)
    return LinesList


def SimpleSplitByDate(split):
    '''
    识别每行，判断是否含有日期，含有则不做处理，不含有则将改行添加到前一行
    '''
    separator = '****'  # 分隔符
    Lines = []  # 读取存储文本每行
    # line = file_read.readline().strip('\n')   #添加ID时需要
    # Lines.append(line)
    for line in split:
        datePattern = '\d{1,2}月(?:\d{1,2}日)?'
        DateMatch = re.search(datePattern, line)
        if DateMatch == None and len(Lines):
            Lines[-1] = Lines[-1] + ' ' + line
        else:
            Lines.append(line)
    return Lines


def MoreSplitByDate(Lines):
    '''
    若为0个，直接添加；
    若为1个，判断该日期是否在行首出现，
        若在，则直接添加，
        不在，则通过日期出现的首位置分隔文本，判断前一行是否为间隔符，
            不是则将日期前文本添加到前一行，日期后文本成为独立一行
            是则将日期前文本独立成一行，日期后文本成为独立一行；
    若为2个及以上，通过判断每个日期出现位置，分隔成多行。
    '''
    separator = '****'  # 分隔符
    DLines = []
    for line in Lines:
        DatePattern = '(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?[和及、])+(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?)|(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)|(?:(?<!\~)(?:\d{4}年)?\d{1,2}月(?:\d{1,2}日)?(?:上午|中午|晚)?(?!\~))'
        # SeveralDatePattern = '(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?[和及、])+(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?)'
        dateS = len(re.findall(DatePattern, line))
        if dateS == 0:
            DLines.append(line)
        elif dateS == 1:
            Pattern = '^(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?'
            if re.search(Pattern, line):  # 查看是否在行首出现月日
                DLines.append(line)
            else:
                Match = re.search(DatePattern, line)
                s = Match.start()
                if DLines[-1] == separator:  # 判断上一句是否为间隔符
                    DLines.append(line[:s])
                else:
                    DLines[-1] = DLines[-1] + ' ' + line[:s]  ##添加到上一句末尾
                DLines.append(line[s:])
        else:
            DateMatch = re.finditer(DatePattern, line)
            lastdatesep = 0
            Flag = True  # 文本以日期开头，则Flag为False
            for date in DateMatch:
                datesep = date.start()
                if datesep != 0:
                    if lastdatesep == 0 and Flag:
                        # if checkdata(DLines[-1]):  #检查上一句是否有日期
                        if DLines[-1] == separator:  # 判断上一句是否为间隔符
                            DLines.append(line[lastdatesep:datesep])
                        else:
                            DLines[-1] = DLines[-1] + ' ' + line[lastdatesep:datesep]  # 添加到上一句末尾
                        lastdatesep = datesep
                    else:
                        DLines.append(line[lastdatesep:datesep])
                        lastdatesep = datesep
                        Flag = True
                else:
                    Flag = False
            DLines.append(line[lastdatesep:])
    return DLines


def MoreSplitByDate_new(Lines):
    '''
    若为0个，直接添加；
    若为1个，判断该日期是否在行首出现，
        若在，则直接添加，
        不在，则通过日期出现的首位置分隔文本，判断前一行是否为间隔符，
            不是则将日期前文本添加到前一行，日期后文本成为独立一行
            是则将日期前文本独立成一行，日期后文本成为独立一行；
    若为2个及以上，通过判断每个日期出现位置，分隔成多行。
    '''
    separator = '****'  # 分隔符
    DLines = []
    for line in Lines:
        date_pattern = r'(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?[和及、])+(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?)|(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)|(?:(?<!\~)(?:\d{4}年)?\d{1,2}月(?:\d{1,2}日)?(?:上午|中午|晚)?(?!\~))'
        date_count = len(re.findall(date_pattern, line))
        if date_count == 0:
            DLines.append(line)
        elif date_count == 1:
            Pattern = r'^(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?'
            if re.search(Pattern, line):  # 查看是否在行首出现月日
                DLines.append(line)
            else:
                Match = re.search(date_pattern, line)
                s = Match.start()
                if len(DLines) == 0 or DLines[-1] == separator:  # 判断上一句是否为间隔符
                    DLines.append(line[:s])
                else:
                    DLines[-1] = DLines[-1] + ' ' + line[:s]  # 添加到上一句末尾
                DLines.append(line[s:])
        else:
            date_match = re.finditer(date_pattern, line)
            last_date_sep = 0
            flag = True  # 文本以日期开头，则Flag为False
            for date in date_match:
                date_sep = date.start()
                if date_sep != 0:
                    if len(DLines) > 0 and last_date_sep == 0 and flag:
                        # if checkdata(DLines[-1]):  # 检查上一句是否有日期
                        if DLines[-1] == separator:  # 判断上一句是否为间隔符
                            DLines.append(line[last_date_sep:date_sep])
                        else:
                            DLines[-1] = DLines[-1] + ' ' + \
                                line[last_date_sep:date_sep]  # 添加到上一句末尾
                    else:
                        DLines.append(line[last_date_sep:date_sep])
                        flag = True
                    last_date_sep = date_sep
                else:
                    flag = False
            DLines.append(line[last_date_sep:])
    return DLines


def GetToDate(Lines):
    '''
    通过前面的处理，会将日期在句子后面格式的文本错误分隔，例如：
    1月26日早上体感不适 居家休息至1月27日
    分隔成
    1月26日早上体感不适 居家休息至\n1月27日
    因此，我们可以通过识别每行，判断是否只含有日期，将其重新修正为：
    1月26日早上体感不适\n居家休息至1月27日
    '''
    OLines = []
    Pattern = r'^(?:(\d{4})年)?\d{1,2}月\d{1,2}日$'
    Pattern2 = r'(?<=\s)(\S+至)$'
    for var in Lines:
        Match = re.search(Pattern, var)
        if Match:
            Match2 = re.search(Pattern2, OLines[-1])
            if Match2:
                text = Match2.group(1)
                line = text + Match.group()
                OLines[-1] = OLines[-1][:Match2.start()]
                OLines.append(line)
        else:
            OLines.append(var)
    return OLines


###############################STEP 2###################################
# split text
def step2(process_text):
    split = split_by_stop(process_text)  # 以文本中句号分隔文本
    lines = split_by_date(split)  # 以日期分隔文本
    lines = clean_lines(lines)  # 去除分隔后小于等于2个字符的段落
    return lines


########################################################################

def process(orig_line, orig_list):
    '''
    【输入】
    orig_line   原文本行，需要提取信息
    list_old    原列表，需要将新信息放到这个列表的后面

    【输出】
    需要添加的项：交通编号、交通方式、出发点、目的地

    【地点提取规则总结】
    片段格式：["*", "ARG1", start, end]

    连词
    "pos/pku"
    c

    地点
    "ner/ontonotes"
    FAC, GPE, ORG, LOC

    "ner/msra"
    LOCATION

    "ner/pku"
    ns

    "srl"
    ["到", "PRED"] 的后一个 ["*", "ARG1"]
    '''

    allow_ambigious = False
    ambigious_list = ['某', '隔壁']
    verb_list = ['到', '从']
    ontonotes_labels = ['FAC', 'GPE', 'ORG', 'LOC']
    msra_labels = ['LOCATION']
    pku_labels = ['ns']
    analysis = HanLP(orig_line)

    slices = []

    def patch_or_append(seg):
        is_new = True
        for i in range(len(slices)):
            ex_seg = slices[i]
            if ex_seg[2] <= seg[2] and ex_seg[3] >= seg[3]:
                is_new = False
                break
            elif seg[2] <= ex_seg[2] and seg[3] >= ex_seg[3]:
                slices[i] = list(seg)
                is_new = False
                break
        if is_new:
            slices.append(list(seg))

    def merge_slice(orig_slices):
        if len(orig_slices) == 0:
            return orig_slices
        slices_merged = [orig_slices[0]]
        for i in range(1, len(orig_slices)):
            last = slices_merged[-1]
            cur = orig_slices[i]
            if last[3] == cur[2]:
                slices_merged[-1][0] = last[0] + cur[0]
                slices_merged[-1][1] = 'MERGED'
                slices_merged[-1][3] = cur[3]
            else:
                slices_merged.append(cur)
        return slices_merged

    # Initial segments
    for seg in analysis['ner/ontonotes']:
        if seg[1] in ontonotes_labels:
            slices.append(list(seg))

    # Additional segments
    for type in [(msra_labels, 'ner/msra'), (pku_labels, 'ner/pku')]:
        slices = merge_slice(slices)
        for seg in analysis[type[1]]:
            if seg[1] in type[0]:
                patch_or_append(seg)

    # SRL magic
    slices = merge_slice(slices)
    for segs in analysis['srl']:
        last = ['', '', 0, 0]
        for seg in segs:
            if last[0] in verb_list and last[1] == 'PRED' and seg[1].startswith('ARG'):
                if (not allow_ambigious):
                    ambigioius = False
                    for phrase in ambigious_list:
                        if phrase in seg[0]:
                            ambigioius = True
                            break
                    if ambigioius:
                        continue
                patch_or_append(seg)
            last = seg

    # Map index of fine grains to index of characters
    word_index_map = []
    cur = 0
    for word in analysis['tok/fine']:
        word_index_map.append(cur)
        cur += len(word)

    old_slices = slices
    slices = []
    for i in range(len(old_slices)):
        # [start, end)
        seg = old_slices[i]
        start = word_index_map[seg[2]]
        end = len(orig_line) if seg[3] == len(
            word_index_map) else word_index_map[seg[3]]
        seg[2] = start
        seg[3] = end
        slices.append(seg)

    # Get conjuctions
    cur = 0
    conj_index_list = []
    for i in range(len(analysis['tok/fine'])):
        word = analysis['tok/fine'][i]
        pos = analysis['pos/pku'][i]
        if pos == 'c':
            conj_index_list.append(cur)
        cur += len(word)

    # Slice by conjunctions
    old_slices = slices
    slices = []
    new_slices = []
    for seg in old_slices:
        is_split = False
        for conj_index in conj_index_list:
            if seg[2] < conj_index < seg[3] - 1:
                relative_index = conj_index - seg[2]
                left_seg = seg.copy()
                right_seg = seg.copy()
                print("!!", left_seg, right_seg)
                left_seg[0] = seg[0][0:relative_index]
                left_seg[3] = relative_index
                right_seg[0] = seg[0][relative_index + 1:]
                right_seg[2] = relative_index + 1
                new_slices.append(left_seg)
                new_slices.append(right_seg)
                is_split = True
                break
        if not is_split:
            slices.append(seg)
    for seg in new_slices:
        patch_or_append(seg)

    # Fully merge
    old_slices = slices
    slices = []
    for seg in old_slices:
        is_unique = True
        for cmp_seg in old_slices:
            if seg[2] == cmp_seg[2] and seg[3] == cmp_seg[3]:
                continue
            elif cmp_seg[2] <= seg[2] and cmp_seg[3] >= seg[3]:
                is_unique = False
                break
        if is_unique:
            slices.append(seg)

    # Filter slices
    old_slices = slices
    slices = []
    for seg in old_slices:
        # if (len(seg[0]) <= 1) or \
        #         (len(seg[0]) <= 2 and seg[1].startswith('ARG')):
        if (len(seg[0]) <= 1) or \
            (len(seg[0]) <= 2 and seg[1].startswith('ARG')) or \
                regex.search(r'\p{Han}', seg[0]) is None:
            continue
        slices.append(seg)

    # Sort slices
    slices = sorted(slices, key=lambda x: x[2])

    start_loc = '' if len(slices) <= 1 else slices[0][0]
    end_loc = ''
    if len(slices) == 1:
        end_loc = slices[0][0]
    elif len(slices) >= 2:
        end_loc = ';'.join([seg[0] for seg in slices[1:]])

    if len(slices) > 0 and start_loc != '':
        start_index = orig_line.find(start_loc)
        left = orig_line[:start_index]
        if '到' in left:
            end_loc = ';'.join([start_loc, end_loc])
            start_loc = ''

    # 识别交通方式
    tp_type_match = re.search(
        r'(出租车|网约车|动车|私家车|旅游大巴|酒店专车|飞机|公交车|商务车|轮渡|旅游团车|滴滴专车|摩托车|救护车|急救车|海汽快车|高铁|电动车|航班|自驾|私家车|船|列车|骑车|驾车|火车|大巴|摩托三轮车|三轮车|小轿车|车|电火车|滴滴车|海汽专线|班车)',
        orig_line)

    # 识别交通编号
    tp_info_match = re.search(
        r'([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-HJ-NP-Z][A-HJ-NP-Z0-9]{4}[A-HJ-NP-Z0-9挂学警港澳])|(班次号\s?[0-9]+)|(双泰\d+号|黎母号|海棠湾号|南方\d号轮渡|紫荆\d+号|粤海铁3号|黎姆号|海装18号)|(\d+)班车|(?:车次|火车)\s?([GCDZTSPKXLY1-9]\d{1,4})|乘([GCDZTSPKXLY1-9]\d{1,4})次火车|(?:国航|南航|海航|航班|航班号|航班号是)\s?([A-Z\d]{2}\d{3,4})|([A-Z\d]{2}\d{3,4})(?:次航|航班)|(?:乘|由|被)\s?(\d+)(?=路|救护车|急救车|车)',
        orig_line)

    tp_type = tp_type_match.group() if tp_type_match is not None else ''
    tp_info = ''

    if tp_info_match is not None and tp_type_match is not None:
        # 提取车牌号
        if tp_info_match.group(1):
            tp_info = tp_info_match.group(1)

        # 提取班次号
        if tp_info_match.group(2):
            tp_info = tp_info_match.group(2)

        # 提取特殊交通 (双泰\d+号|黎母号|海棠湾号|南方\d号轮渡|紫荆\d+号|粤海铁3号|黎姆号|海装18号)
        if tp_info_match.group(3):
            tp_info = tp_info_match.group(3)

        # 提取班车号
        if tp_info_match.group(4):
            tp_info = tp_info_match.group(4)

        # 提取火车有两种格式
        if tp_info_match.group(5):
            tp_info = tp_info_match.group(5)
        elif tp_info_match.group(6):
            tp_info = tp_info_match.group(6)

        # 提取航班号有两种格式
        if tp_info_match.group(7):
            tp_info = tp_info_match.group(7)
        elif tp_info_match.group(8):
            tp_info = tp_info_match.group(8)

        # 提取特殊交通 例如120
        if tp_info_match.group(9):
            tp_info = tp_info_match.group(9)

    print(slices)
    print([tp_info, tp_type, start_loc, end_loc])

    orig_list.extend([tp_info, tp_type, start_loc, end_loc])
    return orig_list


def process_old(line, list1):
    print("input:", line, list1)
    print("input size:", len(list1))

    # 模型可能无法覆盖到的地点词汇列表
    exception_list = ['机场', '市妇幼保健院', '琼海市博鳌镇幸福海小区家中', '市人民医院', '大小洞天', '天涯区升升超市',
                      '河西路口好芙利蛋糕店', '川味饭店', '海鲜广场', '一心堂药店', '群众街', '川味王餐厅', '羊栏市场', '县人民医院',
                      '省人民医院', '临高县碧桂园金沙滩南享养生会所', '南享养生会所', '省定点医院', '临高县临城镇工会',
                      '海口市美兰区三江农场派出所', '新媒体绿都小区', '美兰机场', '临高县半岛阳光小区', '三亚凤凰机场', '湖北襄阳机场',
                      '明珠商业城百佳汇超市', '君澜三亚湾迎宾馆', '海口琼山大园', '文昌公园', '杏林小区旁瓜菜批发市场',
                      '明珠商业城百佳汇超市', '海安港', '三亚市吉阳区君和君泰小区', '三亚市崖州区南滨佳康宾馆', '中西结合医院',
                      '石碌镇城区', '昌化镇昌城村', '三亚定点留观点海角之旅酒店', '万宁市石梅湾九里三期', '三亚市吉阳区凤凰路南方航空城',
                      '白马井镇海花岛']

    lac_result = lac.run(line)
    concat_list = ['天涯海角景区', '海安', '旧港', '海安', '新港', '八爪鱼超市', '天涯区', '海口市', '板桥', '土福湾',
                   '顺泽福湾', '亚龙湾壹号', '百花谷',
                   '亚龙湾', '红塘湾建国酒店', '国光毫生度假酒店', '海棠区', '南田农场', '广西村', '水蛟村委会大园村',
                   '新城路鲁能三亚湾升升超市', '金中海蓝钻小区', '解放路', '旺豪超市', '南海', '嘉园', '广西自治区',
                   '黄姚古镇', '一心堂', '大广岛药品超市', '钟廖村', '山根镇', '水央屈村', '万宁市', '神州半岛',
                   '君临海小区', '万城镇山泉海小区', '崖州区', '南海', '嘉园', '湛江市', '徐闻', '海安港',
                   '南滨佳康宾馆', '崖州区', '东合逸海郡', '云海台小区', '崖城怡宜佳超市', '吉阳区', '双大山湖湾小区',
                   '星域小区', '神州租车公司', '水蛟村', '武汉市', '武汉市', '汉口站', '广西', '南宁', '北海',
                   '义龙西路汉庭酒店新温泉', '乐东', '千家镇', '只文村', '临城路', '亿佳超市', '君澜三亚湾', '迎宾馆',
                   '儋州市', '光村镇', '文昌市', '月亮城', '新海港', '琼海市', '琼海', '嘉积镇', '天来泉二区', '陵水县',
                   '陵水', '英州镇', '椰田古寨', '兴隆镇', '曼特宁温泉', '神州半岛', '君临海小区', '神州半岛',
                   '新月海岸小区', '临高县', '碧桂园', '金沙滩', '浪琴湾', '儋州市', '光村镇', '解放军总医院',
                   '海南医院', '儋州', '客来湾海鲜店', '新港', '龙华区', '东方市', '涛升国际小区', '涛昇国际小区',
                   '广东', '徐闻海安', '文昌', '东郊椰林', '秀英港', '美兰机场', '石碌镇', '昌化镇', '澄迈县', '福山镇',
                   '昌江县', '昌化镇', '昌城村', '三亚市', '东方市', '八所镇', '海口鹏', '泰兴购物广场',
                   '泰兴购物广场海甸三西路', '文化南路', '金洲大厦', '金廉路广电家园', '美兰区三江农场派出所', '蓝天路',
                   '金廉路广电家园', '秀英区', '锦地翰城', '定安县', '香江丽景小区', '石头公园', '美兰区',
                   '海口市琼山区', '金宇街道', '城西镇', '中沙路', '琼山', '灵山镇', '琼山区', '新大洲大道', '昌洒镇',
                   '文城', '澄迈县', '老城镇', '万达广场', '湘忆亭', '骑楼小吃街', '龙桥镇', '玉良村', '老城开发区',
                   '软件园商业步行街', '石山镇', '美鳌村', '那大镇', '儋州街', '三亚市崖州区', '南滨佳康宾馆',
                   '水蛟村委会', '三亚旺豪超市', '福源小区', '新三得羊羔店', '石碌镇晨鑫饭店', '湖北', '湖北省',
                   '武汉市', '湖南', '衡山', '新疆', '琼山区佳捷精品酒店', '滨江店', '江苏', '无锡', '福架村',
                   '皇马假日游艇度假酒店', '坡博农贸市场', '旺佳旺超市南沙店', '凤凰路南方航空城', '迎宾路', '南山寺']
    # concat_list表示需要进行拼接的地名

    split_list = ['南山寺', '天涯海角景区', '百花谷', '亚龙湾', '威斯汀酒店', '云海台小区', '崖城怡宜佳超市',
                  '万应堂大药房购药', '天涯区', '文昌市月亮城', '千百汇超市', '乐东县人民医院', '三亚市人民医院',
                  '君澜三亚湾迎宾馆', '云海台', '崖城怡宜佳超市', '海安新港', '杏林小区旁瓜菜批发市场', '文昌公园',
                  '杏林小区旁瓜菜批发市场', '明珠商业城百佳汇超市', '一心堂', '大广岛药品超市', '南滨农贸市场', '南海嘉园',
                  '三亚市吉阳区', '新风街', '三亚湾', '凤翔路', '中山南路', '老城四季康城', '绿水湾小区', '天涯海角',
                  '玫瑰谷', '福山骏群超市', '那大镇兰洋路', '中兴大道']
    # split_list表示需要用分号隔开的地名

    location_list = []
    seg_list = lac_result[0]
    pos_list = lac_result[1]
    for i in range(len(seg_list)):
        if pos_list[i] == "LOC" or pos_list[i] == "ORG":
            location_list.append(seg_list[i])
    line = line.replace(' ', '')

    tp_type_match = re.search(
        r'(出租车|网约车|动车|私家车|旅游大巴|酒店专车|飞机|公交车|商务车|轮渡|旅游团车|滴滴专车|摩托车|救护车|急救车|海汽快车|高铁|电动车|航班|自驾|私家车|船|列车|骑车|驾车|火车|大巴|摩托三轮车|三轮车|小轿车|车|电火车|滴滴车|海汽专线|班车)',
        line)  # 得到交通工具
    tp_info_match = re.search(
        r'([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领][A-HJ-NP-Z][A-HJ-NP-Z0-9]{4}[A-HJ-NP-Z0-9挂学警港澳])|(班次号\s?[0-9]+)|(双泰\d+号|黎母号|海棠湾号|南方\d号轮渡|紫荆\d+号|粤海铁3号|黎姆号|海装18号)|(\d+)班车|(?:车次|火车)\s?([GCDZTSPKXLY1-9]\d{1,4})|乘([GCDZTSPKXLY1-9]\d{1,4})次火车|(?:国航|南航|海航|航班|航班号|航班号是)\s?([A-Z\d]{2}\d{3,4})|([A-Z\d]{2}\d{3,4})(?:次航|航班)|(?:乘|由|被)\s?(\d+)(?=路|救护车|急救车|车)',
        line)

    start = ''
    end = ''

    for phrase in exception_list:
        if phrase in line:
            end = phrase
            break

    if len(location_list) == 2:
        # 若为地点二元组
        if location_list[0] in concat_list and location_list[1]:
            end = location_list[0] + location_list[1]
            start = ''  # 如果识别到的两个地名，都是在my_list里面，就相当于原先只有一个地名，那么被切分成两个地名，
            # 所以我们需要把这两个地名合起来作为终点，终点就相当于两者的拼接
        else:
            start = location_list[0]
            end = location_list[1]  # 否则的话，就是这两个地点不是由一个地点拆分的，那么我们就
            # 将第一个作为起点，然后将第二个起点作为终点
        if location_list[0] in split_list and location_list[1] in split_list:
            end = location_list[0] + ';' + location_list[1]
            start = ''  # 如果识别到的两个地点都是在my_list里面的话，那么意思就是两个地点都为终点，所以我们需要将这两个地点存入终点中
            # 并且用分号隔开
    elif len(location_list) == 3:
        # 若为地点三元组（与二元组逻辑相似）
        if location_list[0] in concat_list and location_list[1] in concat_list and location_list[2] in concat_list:
            end = location_list[0] + location_list[1] + location_list[2]
            start = ''
        elif location_list[0] in split_list and location_list[1] in concat_list and location_list[2] in concat_list:
            end = location_list[1] + location_list[2]
            start = location_list[0] + ';' + end
        elif location_list[0] in concat_list and location_list[1] in concat_list:
            start = location_list[0] + location_list[1]
            end = location_list[2]
        elif location_list[1] in concat_list and location_list[2] in concat_list:
            start = location_list[0]
            end = location_list[1] + location_list[2]

    if tp_info_match is not None and tp_type_match:  # 判断是否坐的飞机，因为只有飞机有航班号
        # 提取车牌号
        if tp_info_match.group(1):
            t = tp_info_match.group(1)

        # 提取班次号
        if tp_info_match.group(2):
            t = tp_info_match.group(2)

        # 提取特殊交通 (双泰\d+号|黎母号|海棠湾号|南方\d号轮渡|紫荆\d+号|粤海铁3号|黎姆号|海装18号)
        if tp_info_match.group(3):
            t = tp_info_match.group(3)

        # 提取班车号
        if tp_info_match.group(4):
            t = tp_info_match.group(4)

        # 提取火车有两种格式
        if tp_info_match.group(5):
            t = tp_info_match.group(5)
        elif tp_info_match.group(6):
            t = tp_info_match.group(6)

        # 提取航班号有两种格式
        if tp_info_match.group(7):
            t = tp_info_match.group(7)
        elif tp_info_match.group(8):
            t = tp_info_match.group(8)

        # 提取特殊交通 例如120
        if tp_info_match.group(9):
            t = tp_info_match.group(9)

        dict_y = [t, tp_type_match.group(), start, end]
    elif tp_type_match:
        dict_y = ['', tp_type_match.group(), start, end]
    elif tp_info_match:
        dict_y = [tp_info_match.group(), '', start, end]
    else:
        dict_y = ['', '', start, end]

    list1.extend(dict_y)

    print("output:", list1)

    return list1


###############################STEP 3###################################
# get Entity
def step3(line_text):
    sheet = []
    sheet.append(
        ['原文', '病例编号', '性别', '年龄', '常住地', '关系', '日期', '句子', '交通编号', '交通方式', '出发点',
         '目的地'])

    s3 = []  # 将数据用句号进行切割
    s4 = []  # 将数据用句号进行切割的数据
    list1 = []
    s1 = line_text

    data_list = []  # 将每个文本的信息存储到列表中
    for i in s1:  # 从对不规则符号经过处理之后的数据中取值，每次取出的是excel表中每一行存储的文本数据
        flag = False
        routes = []
        nums = re.search(r'第\d+', i)
        if nums:
            num = nums.group()
        else:
            num = ''
        sex = re.search(r'[男|女]', i)  # 正则得到性别
        ages = re.search(r'\d+[\u5c81]', i)  # 正则得到年龄
        citys = re.search(r'常住\w+', i)
        if citys:
            city = citys.group()
        else:
            city = ''
        relation1 = re.search(r'(\d+号\、?\w+病例朋友)', i)
        relation2 = re.search(
            r'(系\d[\w+\、?]+)|(与第\w+[\、?\w+]+)(亲属关系|亲友关系|同一小区|一起乘网约车)', i)
        if relation2 and relation1:
            s = relation2.group() + ',' + relation1.group()
            flag = True
        if relation2:
            if flag:
                relation = s
            else:
                relation = relation2.group()
        else:
            relation = ''
        s3 = (i.split('。'))  # 文本中的数据做句号切割，因为每个文本中包含多句话，而每句话中包含多个需要的信息
        for j in s3:  # 取出来的是每句话
            if j == '\n' or j == '':
                continue
            else:
                time = re.findall(r'\d+月\d+日', j)  # 正则得到时间
                if time:
                    for t in time:
                        list1 = []
                        dict_y = []
                        sentence = j.split(t)[1]
                        times = re.search(r'\d+月\d+日', sentence)
                        sentences = re.findall(
                            r'(乘\w+|转运至\w+|去\w+|到\w+|到达\w+|从\w+|在\w+|进入\w+|返回\w+|抵达\w+|转至\w+|入住\w+|前往\w+|送至\w+|接往\w+|转至\w+|乘坐\w+)',
                            sentence)
                        if times and len(sentences) > 1:
                            list1.extend([i, num, sex.group() if sex else '', ages.group(
                            ) if ages else '', city, relation, t, sentence.split('，')[0]])
                            if sentence:
                                sheet.append(process(sentence, list1))
                        elif 1 < len(sentences):
                            for z in range(len(sentences)):
                                list1 = []
                                list1.extend([i, num, sex.group() if sex is not None else '',
                                              ages.group() if ages is not None else '', city, relation, t, sentence])
                                if sentence[z]:
                                    rest = process(sentences[z], list1).copy()
                                    sheet.append(rest)
                        else:
                            list1.extend([i, num, sex.group() if sex is not None else '',
                                          ages.group() if ages is not None else '', city, relation, t, sentence])
                            if sentence:
                                try:
                                    sheet.append(process(sentence, list1))
                                except:
                                    print(sentence)
                                    print(list1)
                                    r = process(sentence, list1)
                                    print(1, r)
                else:
                    sheet.append(
                        [i, num, sex.group() if sex is not None else '', ages.group() if ages is not None else '', city,
                         relation, '', '', '', '', '', ''])
    return tuple(sheet)


########################################################################

def singleRequest(location="中山大学", limit_city=""):
    """
    单次请求处理
    """
    if location == '':
        return ''

    params = {
        "ak": "ie5WA3RCp4pDqO44jaBluGqYsDClBkdq",
        "region": "null" if limit_city == "" else limit_city,
        "output": "json",
        "query": location,
    }
    query_url = f"https://api.map.baidu.com/place/v2/suggestion?{urlencode(params)}"
    res = requests.get(query_url).json()
    lat = "None"
    lng = "None"
    name = location
    try:
        lat = res["result"][0]["location"]["lat"]
        lng = res["result"][0]["location"]["lng"]
        name = res["result"][0]["name"]
    except:
        pass
    res = [lat, lng, name]
    return res


def allin(list_data):
    # 链接到航班数据库，host为本地主机，端口是3306，用户是root，密码是Scott624，连接上的数据库为flight，编码是utf8
    link = sqlite3.connect('transport.db')

    # 获取游标对象，游标是系统为用户开设的一个数据缓冲区，存放SQL语句的执行结果，结果是零条、一条或多条数据
    data1 = link.cursor()
    data2 = link.cursor()
    data3 = link.cursor()
    data4 = link.cursor()
    write1 = link.cursor()
    write2 = link.cursor()

    ret = ['', '', '']

    # # 获取excel文档的数据表
    # ws = wb['数据']
    #
    # # 命名要存放数据的列
    # ws['N1'] = '机场/火车站点'
    # ws['O1'] = '火车站序'
    # ws['P1'] = '出发时间'
    # ws['Q1'] = '到达时间'
    # # 获取excel数据表的，第I列单元格
    # colI = ws['J']
    # # 获取I列单元格的最大长度
    # _max = len(colI)
    # # 因为第一行数据是不需要进行处理的，所以处理数据要从第二行开始
    # x = 2
    if list_data[9] == '航班':

        # 使用sql语句查询数据库。EXECUTE语句可以执行存放SQL语句的字符串变量，或直接执行SQL语句字符串。需要查询的数据由.format来填充，StartStation.value是出发地，ArriveStation.value[0:2]取目的地的前两个字为查询输入字段，_cell.value是航班号

        if list_data[10] != '' and list_data[11] != '':
            data1.execute(
                f"""select StartDrome,ArriveDrome,StartTime,ArriveTime from tbl_air
                where startCity = (select a.Abbreviation from tbl_icao as a where a.Address_cn='{list_data[10]}')
                and lastCity=(select b.Abbreviation from tbl_icao as b where b.Address_cn='{list_data[11]}')
                and AirlineCode='{list_data[8]}';""")

            # fetchall使用游标从查询中检索
            All = data1.fetchall()
            # 如果All中的数据不为空
            if All:
                # 首先利用sql语句，将数据写入一个新表中。_cell.value是航班号，All[0][0]是出发机场，All[0][1]是降落机场，All[0][2]是起飞时间，All[0][3]是降落时间。同时要在起飞机场、降落机场和航班号都不同的情况下才写入到新的表中
                write1.execute(
                    f"""insert or ignore into tbl_airnum (Code,Start, Arrive, StartTime, ArriveTime)
                    values ('{list_data[8]}','{All[0][0]}','{All[0][1]}', '{All[0][2]}','{All[0][3]}');""")
                # 表示允许写入
                link.commit()

                ret = [All[0][0] + '|' + All[0][1], All[0][2], All[0][3]]

    else:
        # 用于存放列车经过的站点
        Save_station = ""
        # 用于存放列车经过站点的站序
        Save_S_No = ""
        # 用于存放列车的进站时间
        Save_A_Time = ""
        # 用于存放列车的出站时间
        Save_D_Time = ""

        if list_data[8] != '':
            # 获取该行第K、L列的数据。第K列的数据是出发地，第L列的数据为目的地

            # 利用sql语句进行查询。通过列车号、出发站和到达站查询途径站点、站序、进站时间和出站时间。_cell.value是列车号，Train_Start.value是出发站，Train_Arrive.value是到达站
            data2.execute(
                f"""select Station,S_No,A_Time,D_Time from tbl_train where ID = '{list_data[8]}'
                and S_No >= (select S_No from tbl_train where ID = '{list_data[8]}'
                and Station like '{list_data[10]}%' limit 1) and S_No <= (select S_No from tbl_train where ID = '{list_data[8]}'
                and Station like '{list_data[11]}%' limit 1) order by S_No;""")

            # result中的数据分别是站名、站序、到达时间和出发时间
            result = data2.fetchall()

            # 获取result的最后一个元素的索引下标
            Tlen = len(result) - 1

            for row in result:
                # 将所经过的站点用'|'隔开，并在每一个站后加上一个'站'字
                Save_station = Save_station + row[0] + '站' + '|'
                # 将所经过的站点的站序用'|'隔开
                Save_S_No = Save_S_No + str(row[1]) + '|'
                # 将所经过的站点的进站时间用'|'隔开
                Save_A_Time = Save_A_Time + row[2] + '|'
                # 将所经过的站点的出站时间用'|'隔开
                Save_D_Time = Save_D_Time + row[3] + '|'

            # 将所有需要储存的数据的最后一个竖线去除
            Save_station = Save_station[0:len(Save_station) - 1]
            Save_S_No = Save_S_No[0:len(Save_S_No) - 1]
            Save_A_Time = Save_A_Time[0:len(Save_A_Time) - 1]
            Save_D_Time = Save_D_Time[0:len(Save_D_Time) - 1]

            # 如果result中的存有数据
            if result:
                # 将列车号、出发站的站序和到达站的站序储存在新表中。_cell.value是列车号，result[0][1]是出发站的站序，result[Tlen][1]是到达站的站序
                write2.execute(
                    f"""insert or ignore into tbl_trainum (Code,startNum,arriveNum)
                    values ('{list_data[8]}','{result[0][1]}','{result[Tlen][1]}');""")
                link.commit()

        ret = [Save_S_No, Save_A_Time, Save_D_Time]

    # 最后关闭连接
    data1.close()
    data2.close()
    data3.close()
    data4.close()
    link.close()

    return ret
    # 从第九列的第二行到最后一行，读取数据，读取的数据是列车号或者航班号


###############################STEP 4###################################
def step4(step3_result):
    needs = step3_result[1:]
    title = step3_result[0]
    results = []
    title.extend(['出发地标准名称', '出发地经纬度', '目的地标准名称',
                  '目的地经纬度', '航班/站点信息', '出发时间', '到达时间'])
    for data in needs:
        res = None

        if data[9] != '' and data[8] != '':
            res = allin(data)

        if res and '' not in res:
            start = singleRequest(location=res[0].split('|')[0])
            end = singleRequest(location=res[0].split('|')[1])
            data.extend([start[2], ','.join([str(i) for i in start[:2]]), end[2], ','.join(
                [str(i) for i in end[:2]]), res[0], res[1], res[2]])
        else:
            for pos in data[-2:]:
                if pos == '':
                    data.extend(['', ''])
                else:
                    r = singleRequest(location=pos)
                    r = ['', '', '']
                    data.extend([r[2], ','.join([str(i) for i in r[:2]])])
            data.extend(['', '', ''])

        results.append(data)

    final = pd.DataFrame(results, columns=title)
    return final


def step4_fake(step3_result):
    needs = step3_result[1:]
    title = step3_result[0]
    results = []
    title.extend(['出发地标准名称', '出发地经纬度', '目的地标准名称',
                  '目的地经纬度', '航班/站点信息', '出发时间', '到达时间'])
    for data in needs:
        res = None

        if res and '' not in res:
            start = singleRequest(location=res[0].split('|')[0])
            end = singleRequest(location=res[0].split('|')[1])
            data.extend([start[2], ','.join([str(i) for i in start[:2]]), end[2], ','.join(
                [str(i) for i in end[:2]]), res[0], res[1], res[2]])
        else:
            for pos in data[-2:]:
                if pos == '':
                    data.extend(['', ''])
                else:
                    r = ['', '', '']
                    data.extend([r[2], ','.join([str(i) for i in r[:2]])])
            data.extend(['', '', ''])

        results.append(data)

    final = pd.DataFrame(results, columns=title)
    return final


if __name__ == '__main__':
    input_df = pd.read_excel(
        '/Users/lihaojun/Downloads/Copy of CRF_Hainan.xlsx')
    text_list = input_df['additional_information'].to_numpy()[:32]

    frames = []
    for ori_text in text_list:
        text = step1(ori_text)
        process_data = step2(text[0])
        print(process_data)
        step3_data = step3(process_data)
        step4_data = step4_fake(step3_data)
        frames.append(step4_data)
        # print(step4_data.to_string())

    df = pd.concat(frames)
    df.to_csv("out_ernie.csv", sep=',', encoding='utf-8')
