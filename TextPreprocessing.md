# 文本预处理

## 一、概要说明

### 1、摘要

* 文本语料在输送给模型前一般需要一系列的预处理工作, 才能符合模型输入的要求, 如: 将文本转化成模型需要的张量, 规范张量的尺寸等, 而且科学的文本预处理环节还将有效指导模型超参数的选择, 提升模型的评估指标.

### 2.项目需求

* 提取excel文档中数据
* 去除文本中空白行数据
* 统一字符串(包括汉字、数字、字母、标点符号)为半角
* 将小于一岁的年龄标准化表示(年龄=月数/12)
* 去除或修改指定无用的标点符号和干扰文字
* 语义近似的词语替换为统一词语
* 补全不完整日期的月份
* 表示时间的语段格式标准化统一
* 模糊时间表示具体化(可选)
* 查找文本中列车号和航班号,判断是否正确
* 关键词文本分隔：

    通过句号分隔文本

    通过日期分隔文本

### 3.开发环境及工具

* python3.80
* OpenPyXl模块
* Re模块
* datetime模块
* os模块
* visual studio code

## 二、操作流程

### 1、提取excel文档中数据

> 我们拿到的初始数据为exel文本格式，需要我们将其中的指定有用数据提取处理。  

```python
lines = []
wb = xl.load_workbook(file_path)
sheet = wb['Sheet1']
for row in range(2, sheet.max_row + 1):
    line = []
    #读取每段数据ID
    ID =  str(sheet.cell(row,6).value) + '_' + str(sheet.cell(row,5).value) + '_' + str(row-1)
    #读取每段数据原文信息
    orig_text = sheet.cell(row, 22)
    text = 文本预处理函数(orig_text)
    line.append(ID)
    line.append(orig_text)
    line.append(text)
    lines.append(line)
```

我们使用二维列表存储提取的数据，使用for循环提取excel每行的数据，创建添加到二维列表。创建的二维列表如下：

```python
lines = [[ID_1,orig_text_1,text_1],
         [ID_2,orig_text_2,text_2],
         ...
         [ID_n,orig_text_n],text_n]
```

其中的相关参数说明如下：  

* lines：表示存储所有行数据的二维列表名称  
* line：表示存储各单行数据的一维列表名称  
* ID：每段数据的唯一关键词  
* orig_text：每段未经过文本预处理的初始数据  
* text：每段文本预处理后的数据

### 2、去除文本中空白行数据

>我们从excel文档中提取的文本含有我们不想要的空行，需要我们将其去除。

```python
def clearBlankLine(text):
    clear_lines = ""
    for line in text:
        if line != '\n':
            clear_lines+=line
    return clear_lines
```

我们使用for循环识别判断每段数据中是否存在换行符，以保留有效数据。  
传入函数的text为每段文本需要处理的数据，返回为去除空行的数据。  
例如我们需要处理的源文本为：

```python
text = '第3号确诊病例（三亚）

男，27岁，常住地湖北武汉。1月17日身体不适→1月18日乘JD5628次航班从武汉到三亚，转乘出租车到三亚市天涯区嘉濠旅租住处→1月20日13:30左右乘出租车到三亚市人民医院就诊，1月22日确诊，2月2日出院。'
```

调用以上函数返回结果为：

```python
第3号确诊病例（三亚）男，27岁，常住地湖北武汉。1月17日身体不适→1月18日乘JD5628次航班从武汉到三亚，转乘出租车到三亚市天涯区嘉濠旅租住处→1月20日13:30左右乘出租车到三亚市人民医院就诊，1月22日确诊，2月2日出院。
```

### 3、统一字符串(包括汉字、数字、字母、标点符号)为半角

>通过分析文本，我们发现文本中含有大量全角和半角的标点符号、字母等字符(全角情况下输入一个字符就会占用两个字符，半角情况下输入一个字符只占用一个字符)，格式并不统一。我们可以通过正则表达式将其识别并处理。

```python
def Q2B(uchar):
    """单个字符 全角转半角"""
    inside_code = ord(uchar)
    if inside_code == 12288:  # 全角空格直接转换
        inside_code = 32
    elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
        inside_code -= 65248
    return chr(inside_code)

def stringQ2B(ustring):
    """把字符串全角转半角"""
    return "".join([Q2B(uchar) for uchar in ustring])
```

我们调用ord()函数返回对应的ASCII数值，或者Unicode数值，我们通过Unicode编码规则，判断相应编码的字符是否为全角，并替换全角字符为半角字符。

* 全角即：Double Byte Character，简称DBC  
* 半角即：Single Byte Character，简称SBC  

相关具体规则为：  

* 全角字符unicode编码从65281~65374 （十六进制 0xFF01 ~ 0xFF5E）  
* 半角字符unicode编码从33~126 （十六进制 0x21~ 0x7E）  
* 空格比较特殊，全角为 12288（0x3000），半角为 32（0x20）  
* 而且除空格外，全角/半角按unicode编码排序在顺序上是对应的（半角 + 65248 = 全角）  

相关函数介绍:  

* chr()函数用一个范围在range（256）内的（就是0～255）整数作参数，返回一个对应的字符。  
* unichr()跟它一样，只不过返回的是Unicode字符。
* ord()函数是chr()函数或unichr()函数的配对函数，它以一个字符（长度为1的字符串）作为参数，返回对应的ASCII数值，或者Unicode数值。

例如我们需要处理的源文本为：

```python
text = 'Hainan_Lingao_85罗XX，女，49岁，常住到湖北武汉武昌区，与31日确诊患者孙XX为母女关系，１月21日从湖北省武汉市乘坐飞机(海航ＨＵ7581)到海南。现住临高县博厚镇碧桂园浪琴湾,否认既往疾病史。１月28日出现发热伴咳嗽、胸闷,未到过医疗机构治疗。２天后不见好转,于１月30日到县医院就诊,拟“肺炎”收治入院。１月30日第一次采样送检,为阴性；1月31日第二次采样送检,为阳性,1月31日转院至省人民医院。已登记密切接触者６人,３人现住到南享会所,３人居住到碧桂园浪琴湾。以上密切接触者均已采取居家隔离观察措施。'
```

调用以上函数返回结果为：

```python
Hainan_Lingao_85罗XX,女,49岁,常住到湖北武汉武昌区,与31日确诊患者孙XX为母女关系,1月21日从湖北省武汉市乘坐飞机(海航HU7581)到海南。现住临高县博厚镇碧桂园浪琴湾,否认既往疾病史。1月28日出现发热伴咳嗽、胸闷,未到过医疗机构治疗。2天后不见好转,于1月30日到县医院就诊,拟“肺炎”收治入院。1月30日第一次采样送检,为阴性;1月31日第二次采样送检,为阳性,1月31日转院至省人民医院。已登记密切接触者6人,3人现住到南享会所,3人居住到碧桂园浪琴湾。以上密切接触者均已采取居家隔离观察措施。
```

### 4、将小于一岁的年龄标准化表示(年龄=月数/12)

>通过分析文本，我们发现文本中小于一岁的年龄表示为 n个月，不同的年龄表示对于年龄识别有更多要求，所以我们可以通过将所有年龄统一格式标准化，来简化识别难度。

```python
def StandardAge(text):
    """将小于一岁的年龄标准化表示(年龄=月数/12)"""
    AgePattern = '(?:,)(\d{1,2})个月(?:,)'
    match = re.search(AgePattern,text)
    if match:
        s = match.start()
        e = match.end()
        age = round(int(match.group(1))/12, 2)
        text = text[:s]+str(age)+'岁'+text[e:]
    return text
```

其中正则表达式中相关参数说明如下：

* '(?:,)(\d{1,2})个月(?:,)' :识别前后带有逗号(半角)的n个月格式，例如：女，3个月，常住地湖北孝感市。 防止识别到具有其他含有的n个月表示。

例如我们需要处理的源文本为：

```python
text = '第37号确诊病例（海口）女,3个月,近3个月常住地湖北孝感市。'
```

调用以上函数返回结果为：

```python
第37号确诊病例（海口）女0.25岁近3个月常住地湖北孝感市。
```

### 5、去除或修改指定无用的标点符号和干扰文字

>通过分析文本，我们发现文本中含有许多影响文本阅读或去除修改后对文本无影响的标点符号和文字，我们可以通过正则表达式将其识别并处理。

```python
def clean_data(text):
    '''删除不表示时间的冒号,删除'--',删除空格'''
    text = re.sub(r'(?<=\D):|--|\s','',text) 
    '''删除序列标号 例如: 1. 3. 1、 2、'''
    text = re.sub(r'[1-9]\d*(?:\.(?:(?=\d月)|(?!\d))|\、)','',text) 
    '''替换逗号,后括号 为 空格'''
    text = re.sub('[,)]',' ',text) 
    '''删除前括号'''
    text = re.sub('\s?\(','',text) 
    '''替换箭头(->或→)为句号'''
    text = re.sub(r'->|→','。',text) 
    '''替换分号为句号'''
    text = re.sub(r';','。',text)
    '''删除无用干扰文字'''
    text = re.sub(r'微信关注椰城','',text)
    '''将因前面处理后生成的句号而导致两个连续句号产生,替换为一个句号'''
    text = re.sub(r'。。','。',text)
    return text
```

其中正则表达式中相关参数说明如下：  

* (?<=\D): 表示查找不表示时间格式(例如:12:10)的冒号  
* \d\. 表示数字加点  例如:1. 2.  
* \s   表示空格，包括空格、换行、tab缩进等所有的空白  

例如我们需要处理的源文本为：

```python
text = '男,61岁,常住重庆市万盛区。1.1月11日至16日,参加由逍遥旅行社组团到泰国旅游。乘坐的航班号是SL1439和SL8438,经核查该航班有两例确诊病例,分别是海南省确诊病例49号和海南省确诊病例69号。航班号SL1439,在1月11日13:58分从(微信关注椰城)美兰机场起飞->14:58分到达泰国,航班号SL8438,在1月16日22:56分从:泰国起飞,1月17日凌晨01:46分到达海口美兰机场→文昌团12人:黄某雄、黄某元、任某珍、方某秀、朱某勇、唐某芳、王某玫、方某容、罗某贵、张某才、余某兴、李某群，其中余某兴、李某群已回重庆。2.1月17日至18日,确诊病例自述步行到文昌市文城第一市场、杏林小区旁瓜菜批发市场、明珠商业城百佳汇超市购物和到文昌公园散步。'
```

调用以上函数返回结果为：

```python
男61岁常住重庆市万盛区。1月11日至16日参加由逍遥旅行社组团到泰国旅游。乘坐的航班号是SL1439和SL8438经核查该航班有两例确诊病例分别是海南省确诊病例49号和海南省确诊病例69号。航班号SL1439在1月11日13:58分从美兰机场起飞。14:58分到达泰国航班号SL8438在1月16日22:56分从泰国起飞1月17日凌晨01:46分到达海口美兰机场。文昌团12人黄某雄、黄某元、任某珍、方某秀、朱某勇、唐某芳、王某玫、方某容、罗某贵、张某才、余某兴、李某群，其中余某兴、李某群已回重庆。1月17日至18日确诊病例自述步行到文昌市文城第一市场、杏林小区旁瓜菜批发市场、明珠商业城百佳汇超市购物和到文昌公园散步。
```

### 6、语义近似的词语替换为统一词语

>通过分析文本，我们发现统一修改一部分词语，不会影响文本含义，并且可以提高识别效率，简化文本结构。

```python
def clean_data(text):
    def goTo(text):
        '''修改替换文本中不以句号结束的行动动词，统一为‘到’'''
        Pattern = '(?:抵达|返回|进入|在|回|赴|回到|前往|去到|飞往)(?!\。)'
        Match=re.search(Pattern,text)
        while Match:
            s = Match.start()
            e = Match.end()    
            text = text[:s] + '到' + text[e:]
            Match=re.search(Pattern,text)
    text = goTo(text)
    '''替换文本中表示乘坐的动词，统一为‘乘’'''
    text = re.sub(r'乘坐|搭乘|转乘|坐','乘',text)
    '''统一文本中火车的两种表示方式(动车或列车)为‘火车’'''
    text = re.sub(r'动车|列车','火车',text)
    '''替换文本中‘飞机’为‘航班’'''
    text = re.sub(r'飞机','航班',text)
    return text
```

为什么只替换文本中不以句号结束的行动动词:  
防止替换了修改后会影响语义通顺的行动动词。例如:在11:20已经到达。

例如我们需要处理的源文本为：

```python
text = '''1月24日15:20左右前往亚特兰蒂斯水族馆游玩
2019年11月底自齐齐哈尔乘坐飞机抵达三亚
1月21日搭乘FD438次航班从泰国返回三亚
进入海安旧港乘紫荆11号轮渡前往海口
转乘出租车至三亚市天涯区住处
11:00在金鸡岭农贸市场买水果
之后步行回酒店
参加海外国际旅行社组织的赴泰国旅行
乘SL8438次航班从泰国回到海口
1月21日8时45分从武汉乘飞机到海口国航CA8235经济舱25排过道
2020年1月15日乘C7317次列车1车厢于从美兰到三亚
1月25日乘C7419次动车6车厢到琼海
'''
```

调用以上函数返回结果为：

```python
1月24日15:20左右到亚特兰蒂斯水族馆游玩
2019年11月底自齐齐哈尔乘飞机到三亚
1月21日乘FD438次航班从泰国到三亚
到海安旧港乘紫荆11号轮渡到海口
乘出租车到三亚市天涯区住处
11:00到金鸡岭农贸市场买水果
之后步行到酒店
参加海外国际旅行社组织的到泰国旅行
乘SL8438次航班从泰国到海口
1月21日8时45分从武汉乘航班到海口国航CA8235经济舱25排过道
2020年1月15日乘C7317次火车1车厢于从美兰到三亚
1月25日乘C7419次火车6车厢到琼海
```

### 7、补全不完整日期的月份

>通过分析文本，我们发现文本中含有大量不指代月份的日期，这对日期的提取会有很大的影响，我们可以通过前后文最近完整日期查找，来补全不完整日期。

```python
def PatchMonth(text):  
    '''补全月份:将文本中没有月份表示的日期补全为带有月份的日期 '''
    Pattern = '(?<!月|\d)(\d{1,2})日'
    match = re.search(Pattern,text)
    while match:
        day = match.group(1)
        s = match.start()
        e = match.end()
        prevtext = text[:s] #不含月份的日期文本的前文
        nexttext = text[e:] #不含月份的日期文本的后文
        ptime = re.findall('(\d{1,2})月(\d{1,2})日',prevtext)  #查找前文是否含有带有月份的日期
        ntime = re.findall('(\d{1,2})月(\d{1,2})日',nexttext)  #查找后文是否含有带有月份的日期
        if ptime: 
            '''查找到前文距离需补全日期最近的含月日期,比较日期'''   
            if int(ptime[-1][1]) <= int(day): 
                month = ptime[-1][0] + '月'   #前文日期小于等于需补全日期，则需补全的月即为前文的月
            else: month = str(int(ptime[-1][0])+1) + '月'  #前文日期大于需补全日期，则需补全的月即为前文的月+1
            date = month+day + '日'
            text = text[:s]+date+text[e:] 
            match = re.search(Pattern,text)
        elif ntime:
            '''查找到后文距离需补全日期最近的含月日期,比较日期'''
            if int(ntime[0][1]) < int(day):
                month = str(int(ntime[0][0])-1) + '月' #后文日期小于需补全日期，则需补全的月即为后文的月-1
            else: month = ntime[0][0] + '月'  #后文日期大于等于需补全日期，则需补全的月即为后文的月
            date = month+day + '日'
            text = text[:s]+date+text[e:] 
            match = re.search(Pattern,text)
        else: break
    return text
```

查找方式：通过查找不含月份日期文本的前文或后文中含有月份表示的最近日期,比较日的大小,判断补全月份(月份不变或+1或-1)

例如我们需要处理的源文本为：

```python
text = '1月23日到琼山区龙塘镇仁三村委会道本村 25日上午11时从龙塘镇道本村到海口府城开始出车拉客'
```

调用以上函数返回结果为：

```python
1月23日到琼山区龙塘镇仁三村委会道本村 1月25日上午11时从龙塘镇道本村到海口府城开始出车拉客
```

### 8、表示时间的语段格式统一

>通过分析文本，我们发现文本中对时间的表示多种多样，并不规范，语法识别较为困难。

 ***月+日时间有以下格式：***  
    **1月24-31日**  
    **1月22日-2月1日**  
    **2月1日至7日**  
    **1月21日至2月1日**  
    **1月20至24日**
    **1月24日中午-1月26日**
    **2019年12月18日-2020年1月19日上午**
    **2020年1月25日下午至2月1日**
***将以上时间格式统一修改为：***  
**1月2日~3月4日**  
**2019年12月18日~2020年1月19日上午**

```python
def StandardTime(text):
    '''
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
    '''    
    Pattern = '(\d{4}年)?(\d{1,2})月(\d{1,2})日?(上午|中午|下午|晚)?[-至](\d{4}年)?(\d{1,2})[日月](\d{1,2})?日?(上午|中午|下午|晚)?'
    match = re.search(Pattern,text)
    def retut(var):
        '''判断变量是否为None,不为None则返回变量,否则返回空字符'''
        if var:
            return var
        else:return ''
    while match:
        '''提取时间表示中可能含有的信息:'''
        SYear = retut(match.group(1)) #开始年份
        EYear = retut(match.group(5)) #结束年份
        Stime = retut(match.group(4)) #开始某日具体时间(上午|中午|晚)
        Etime = retut(match.group(8)) ##结束某日具体时间(上午|中午|晚)
        s = match.start()
        e = match.end()
        if match.group(7): 
            '''处理含有两个月份的时间段:例如:2月13日至2月16日'''
            #'time = '$'+SYear+match.group(2)+'月'+match.group(3)+'日'+Stime+'~'+EYear+match.group(6)+'月'+match.group(7)+'日'+Etime+'$'
            time = SYear+match.group(2)+'月'+match.group(3)+'日'+Stime+'~'+EYear+match.group(6)+'月'+match.group(7)+'日'+Etime
        else:
            '''处理只含有一个月份的时间段:例如:1月28日至30日'''
            #time = '$'+SYear+match.group(2)+'月'+match.group(3)+'日'+Stime+'~'+EYear+match.group(2)+'月'+match.group(6)+'日'+Etime+'$'
            time = SYear+match.group(2)+'月'+match.group(3)+'日'+Stime+'~'+EYear+match.group(2)+'月'+match.group(6)+'日'+Etime
        text = text[:s] + time + text[e:]
        match = re.search(Pattern,text)
    return text 
```

我们使用Python标准库中Re包(正则表达式)，使用其中search函数，匹配符合正则表达式规则的字符串，并使用while循环依次匹配处理符合规则的字符串。  

其中的相关参数说明如下：  

* Pattern：正则表达式规则,通过此规则可以识别以上几种时间格式  
* re.search：扫描整个字符串并返回第一个成功的匹配，如果没有匹配，就返回一个None  
* match：每次匹配的符合正则表达式规则的字符串  
* s：匹配字符串的开始位置  
* e：匹配字符串的结束位置  
* match.group：在正则表达式中使用(...)分组，而match.group()返回匹配结果中一个或多个group。如果参数值是0,那么返回整个匹配结果的字符串；如果它是[1..99]之间的数字，则返回的是与对应括号组匹配的字符串。  
time：将匹配字符串时间格式统一修改后的结果

例如我们需要处理的源文本为：

```python
text = '''1月24日中午-1月26日入住三亚市天涯区某酒店
2019年12月18日-2020年1月19日上午一直居住陵水碧桂园珊瑚宫殿7栋主要到家周边、偶去家润发超市和附近菜市场买菜 期间曾去三亚市和东方市
1月25日晚-1月27日被安排到兵工大酒店集中隔离观察
2月13日晚至2月16日居家未外出
1月21-23日到三亚一些景点游玩\n1月26日-2月1日隔离观察\n1月25日至2月1日在住处隔离观察未外出\n1月28日至30日居家未外出\n1月20至24日自驾车从武汉出发
2020年1月25日下午至2月1日
2019年12月18日-2020年1月19日上午
1月24日中午-1月26日   
'''
```

调用以上函数返回结果为：

```python
$1月21日~1月23日$到三亚一些景点游玩
$1月26日~2月1日$隔离观察
$1月25日~2月1日$在住处隔离观察未外出
$1月28日~1月30日$居家未外出
$1月20日~1月24日$自驾车从武汉出发
$2020年1月25日下午~2月1日$
$2019年12月18日~2020年1月19日上午$
$1月24日中午~1月26日$
```

***时+分时间有以下格式：***  
**13:30**  
**18时左右**  
**1时30分左右**  
**凌晨2:00左右**  
***将以上时间格式按照分是否为0修改为：***  
**分钟为0的修改为 某时**  
**分钟不为0的修改为 某时某分**  

```python
def timech(text):
    Pattern = '(\d{1,2}):(\d{1,2})'
    match = re.search(Pattern,text)
    while match:
        if match.group(2) == '00':
            hour = match.group(1)
            #time = '$'+hour+'时'+'$'
            time = hour+'时'
        else: 
            hour = match.group(1)
            minute = match.group(2)
            #time = '$'+hour+'时'+minute+'分' + '$'
            time = hour+'时'+minute+'分'
        text = re.sub(Pattern,time,text,count=1)    
        match = re.search(Pattern,text)
    text = re.sub('分分','分',text)
    return text
```

其中的相关参数说明如下：  

* Pattern：正则表达式规则
* match：每次匹配符合正则规则的字符串
* match.group(2)：匹配结果中的分钟表示
* match.group(1)：匹配结果中的时钟表示

例如我们需要处理的源文本为：

```python
text = '1月23日，11:00前往天涯区升升超市购物\n1月24日，12:30前往亚龙湾喜来登酒店入住\n'
```

调用以上函数返回结果为：

```python
1月23日，$11时$前往天涯区升升超市购物
1月24日，$12时30分$前往亚龙湾喜来登酒店入住  
```

### 9、模糊时间表示具体化(可选)

>通过分析文本，我们发现文本中含有少量表示模糊时间的词语，例如当天、当晚，我们可以通过从模糊时间开始向前搜索具体时间的方法，匹配最近时间，以此替换模糊时间。  

我们使用Python标准库中Re包(正则表达式)，使用其中search函数，构建匹配当天、当晚的正则表达式，匹配符合此规则的字符串，并使用while循环依次匹配处理符合规则的字符串。在while循环中，构建匹配在 当天、当晚 前表示时间的正则表达式，使用findall函数匹配符合此规则的所有时间，使用列表存储，列表中最后一项即为离模糊时间最近的具体时间，用此时间替换模糊时间即可。

```python
def ModifyBlurTime(text):
    Pattern = '当天|当晚'
    match = re.search(Pattern,text)
    while match:
        time = re.findall('((?:\d{4}年)?\d{1,2}月\d{1,2}日)(?=.+(?:当天|当晚))',text)
        s = match.start()
        e = match.end()
        if time:
            if match.group() == '当晚':
                text = text[:s]+'$'+str(time[-1])+'晚上'+'$'+text[e:]
            else: text = text[:s]+'$'+str(time[-1])+'$'+text[e:]
            match = re.search(Pattern,text)
        else:break    
    return text
```

其中的相关参数说明如下：  

* Pattern：匹配模糊时间的正则表达式规则  
* match：匹配模糊时间结果  
* time：利用findall函数及规则匹配的时间列表  
* s：模糊时间字符串开始位置  
* e：模糊时间字符串结束位置  
* time[-1]：离模糊时间最近(向前查找)的具体时间  

```python
re.findall(pattern, string, flags=0)
```

findall函数参数讲解：利用pattern规则获取string中所有符合规则的字符串，并以列表的形式返回。flags表示正则表达式的常量。  

例如我们需要处理的源文本为：

```python
text = '2020年1月22日乘MU2527次航班从武汉到三亚→当晚乘私家车到三亚市天涯区水蛟村住处\n2020年1月15日乘MF8315次航班从杭州到海口→当天乘C7317次列车(1车厢)于从美兰到三亚'
```

调用以上函数返回结果为:  

```python
2020年1月22日乘MU2527次航班从武汉到三亚→$2020年1月15日晚上$乘私家车到三亚市天涯区水蛟村住处
2020年1月15日乘MF8315次航班从杭州到海口→$2020年1月15日$乘C7317次列车(1车厢)于从美兰到三亚
```

### 10、查找文本中列车号和航班号,判断是否正确

>通过分析文本，我们发现文本中含有大量列车号和航班号，由于初始文本为人工输入，所以这些车次号可能会出现格式错误，导致后续车次号的识别出现问题。

```python
def IsExistTrainAndFlight(text):
    '''寻找列车号和航班号,通过数据库查找是否存在'''
    TrainPattern = '(?:车次|火车)\s?([GCDZTSPKXLY1-9]\d{1,4})|乘([GCDZTSPKXLY1-9]\d{1,4})次火车'
    FlightPattern = '(?:国航|南航|海航|航班|航班号)\s?([A-Z\d]{2}\d{3,4})|([A-Z\d]{2}\d{3,4})(?:次航班|航班)'
    def LoadDataBase(txts):
        '''加载数据库'''
        f=open(txts, encoding='utf-8')
        txt=[]
        for line in f:
            txt.append(line.strip())
        texts = ';'.join(txt)
        return texts
    #加载列车号数据库
    TrainNumberDataBases = LoadDataBase('train_number.txt')
    #加载航班号数据库
    FlightNumberDataBases = LoadDataBase('flight.txt')
    #寻找
    def find(type,Pattern,DataBase,text):
        '''
        type:编号类型(火车车次或航班)
        Pattern:查找模型(火车模型或航班模型)
        DataBase:加载数据库
        text:需查找的文本
        '''
        for Match in re.finditer(Pattern,text):
            s = Match.start()
            if text[s-1:s] == ' ': #删除航班或火车号前面的空格
                text = text[:s-1]+text[s:]
            if Match.group(1) != None:
                Number = Match.group(1)
            else: Number = Match.group(2)
            if re.search(Number,DataBase) == None:
                print(type+Number+' not Exist')   #查找数据库不存在的列车号和航班号
            # if re.search(Number,DataBase) != None:   查找数据库存在的列车号和航班号
            #     print(type+Number+' Exist')
    find('火车车次:',TrainPattern,TrainNumberDataBases,text)
    find('航班号:',FlightPattern,FlightNumberDataBases,text)
    return text
```

其中正则表达式中相关参数说明如下：

* \s? ：表示此位置含有一个或0个空格

* TrainPattern='(?:车次|火车)\s?([GCDZTSPKXLY1-9]\d{1,4})|乘([GCDZTSPKXLY1-9]\d{1,4})次火车'：分析文本，发现文本中的列车号有 车次XXXX，火车XXXX，乘XXXX次火车 几种格式，通过正则表达式，我们就可以提取关键词附近的列车号

* FlightPattern='(?:国航|南航|海航|航班|航班号)\s?([A-Z\d]{2}\d{3,4})|([A-Z\d]{2}\d{3,4})(?:次航班|航班)'：分析文本，发现文本中的航班号有 国航XXXX，南航XXXX，海航XXXX，航班XXXX，航班号XXXX，XXXX次航班，XXXX航班 几种格式，通过正则表达式，我们就可以提取关键词附近的航班号

**我们可以通过识别车次号，将车次号与数据库中正确车次号进行匹配，判断其是否正确，并返回不正确的车次号。**

例如我们需要处理的源文本为：

```python
text = 
'''
1月16日乘HU7578次航班从武汉到三亚。
1月19日乘CZ3341次航班从武汉到海南乘动车到陵水县土福湾顺泽福湾住处。
乘11时05分至13时30分的航班GJ8921襄阳直飞海口再乘高铁车次C7473美兰至三亚凤凰机场15时33分分发车到三亚。
从武汉乘航班到海口国航CA5444经济舱25排过道。
1月20日从武汉到香港游玩1月27日从香港乘航班到海口KA694航班座位号32排
2020年1月15日乘C4444次火车1车厢于从美兰到三亚。
1月17日乘D8222次火车7车厢到广西北海入住北部湾一号途安海景度假公寓。
'''
```

调用以上函数返回结果为：

```python
火车车次:C4444 not Exist
火车车次:D8222 not Exist
航班号:CA5444 not Exist
```

### 11、关键词文本分隔

```python
def SplitText(file_path,Save_path_ByStop,Save_Path_ultim,Save_Path_textToexcel):
    '''
    按规则分隔文本:
    file_path:需分隔的excel文本路径
    Save_path_ByStop:通过句号为间隔符分隔处理后的text文本路径
    Save_Path_ultim:通过句号和日期分隔处理后的text文本路径
    Save_Path_textToexcel:通过分隔处理后的excel文本路径
    '''
    LInes = [] #纯粹分隔处理后的文本
    SplitByStop(file_path,Save_path_ByStop) #以文本中句号分隔文本
    LInes = SplitByDate(Save_path_ByStop)  #以日期分隔文本
    LInes = ClearLines(LInes)  #去除分隔后小于等于2个字符的段落
    #文本存储为text
    file_write = open(Save_Path_ultim,'w',encoding='UTF-8')
    for var in LInes:
            file_write.writelines(var)
            file_write.writelines('\n')
    file_write.close()
    #文本覆盖纯粹为excel
    textToexcel(Save_Path_ultim,Save_Path_textToexcel)
```

### 11.1 通过句号分隔文本

>通过分析文本，我们发现对每段数据的整段进行处理，会出现大量问题，因此我们可以通过对每段数据文本进行关键词切割，来减轻后续识别处理负担。

```python
def SplitByStop(excelfile_path,Save_textpath_ByStop):
    '''
    以文本中句号分隔文本
    excelfile_path:需识别的excel文档地址
    Save_textpath_ByStop:分隔处理后生成的text文本地址
    '''
    file_write = open(Save_textpath_ByStop,'w',encoding='UTF-8')
    wb = xl.load_workbook(excelfile_path) #加载excel文本
    sheet = wb['Sheet']
    for row in range(2, sheet.max_row + 1):
        orig_text = str(sheet.cell(row, 3).value)
        #添加ID:
        #ID =  str(sheet.cell(row,1).value)
        #file_write.writelines(ID)
        #file_write.writelines('\n')
        point  = re.search('。',orig_text)
        while point:
            s = point.start()
            e = point.end()
            var = orig_text[:s]
            file_write.writelines(var)
            file_write.writelines('\n')
            orig_text = orig_text[e:]
            point  = re.search('。',orig_text)
        file_write.writelines('****\n')
    file_write.close()
```

**我们将前期处理生成的excel文本使用该函数加载，提取每段数据，并依次对每段数据进行句号识别，通过句号分隔文本(将句号替换为换行符\n)，并添加到text文本中。**

调用以上函数返回的部分结果为：

```python
第3号确诊病例三亚 男 27岁 常住地湖北武汉
1月17日身体不适
1月18日乘JD5628次航班从武汉到三亚 乘出租车到三亚市天涯区嘉濠旅租住处
1月20日13时30分左右乘出租车到三亚市人民医院就诊 1月22日确诊 2月2日出院
****
第5号确诊病例三亚 女 27岁 系3号病例朋友 常住地湖北武汉
1月18日乘JD5628次航班从武汉到三亚 乘出租车到三亚市天涯区嘉濠旅租入住 一直未外出
1月20日13时30分左右与3号病例乘出租车到三亚市人民医院就诊 隔离治疗
1月23日确诊 2月7日出院
****
第9号确诊病例三亚 女 65岁 常住地湖北武汉
1月18日1行6人自驾从武汉到海南
1月19日身体不适
1月20日14时左右到海口市板桥市场用餐
自驾到三亚市榆亚路锦轩江南酒店入住 之后到第一市场购物
1月21日自驾到南山寺和天涯海角景区游玩
1月22日到亚龙湾热带公园游玩
1月23日自驾到301医院海南医院发热门诊就诊 隔离治疗
1月24日确诊 2月5日出院
****
第10号确诊病例三亚 男 29岁 常住地湖北武汉
1月21日乘MU2527次航班从武汉到三亚
乘出租车到三亚市海棠区理文索菲特酒店入住
1月22日16时左右乘网约车到301医院海南医院就诊 隔离治疗
1月24日确诊 1月29日出院
****
第11号确诊病例三亚 女 47岁 常住地湖北武汉
1月16日乘HU7578次航班从武汉到三亚
乘出租车到三亚市吉阳区亚龙湾壹号住处
1月19日身体不舒服
1月23日被救护车转到三亚市中心医院就诊
1月24日确诊 2月6日出院
****
第19号确诊病例三亚 女 40岁 常住地湖北武汉
1月19日乘CZ3341次航班从武汉到海南 乘火车到陵水县土福湾顺泽福湾住处
1月20日早乘私家车到英州镇农贸市场买菜
1月21日居家未外出
1月22日身体不适 乘出租车到301医院海南医院就诊
1月23日乘出租车到英州镇农贸市场购物
1月24日乘出租车到301医院海南医院就诊并隔离
1月25日确诊
...
****
第78号确诊病例万宁 男 28岁 系13号、14号、30号、41号、79号病例亲属 常住地湖北武汉
2020年1月20日与家人乘私家车从武汉到海南
1月21日9时50分左右到海安新港安检所乘船号不详 
中午到达海口 到海口骑楼小吃街吃饭
当天乘私家车到三亚市吉阳区某小区
1月22日曾到附近购物、爬山
1月23日~1月25日期间多次自驾到三亚和万宁之间往返 1月25日从万宁接家人到南山寺游玩
当天下午到返程途中某个餐馆用餐
1月26日~1月27日居家未外出
1月28日开车到万宁神州半岛 并被120救护车送至万宁市人民医院就诊
2月3日确诊 目前正到定点医院隔离治疗
****
第84号确诊病例保亭 男 67岁 系85号病例亲属 常住地河南郑州
2019年10月7日与85号病例乘JD5726次航班从郑州到三亚
乘车到保亭县某小区住处
期间未曾离开保亭 平时除购买生活用品外以居家为主
2020年1月25日与85号病例一起参加小区组织的拍集体照活动
2月2日乘私家车到三亚市定点医院隔离治疗
2月4日确诊 目前正到定点医院隔离治疗
...

```

### 11.2 通过日期分隔文本

>通过分析句号分隔后的文本，我们发现text文本中，每组数据仍存在一些问题，例如：应属于一行的文本信息，因为前期句号处理不当，而被分开；句号分隔后的数据中，还有每行文本含有多个日期行程的情况,因此，我们还需要通过识别日期来再次分隔文本。

```python
def SplitByDate(file_path):
    '''以日期分隔文本'''  
    LinesList = []
    LinesList = SimpleSplitByDate(file_path)
    LinesList = MoreSplitByDate(LinesList)
    #CDLines = SplitByCTNDate(DLines)
    LinesList = GetToDate(LinesList)
    return LinesList
```

```python
def SimpleSplitByDate(file_path):
    '''
    识别每行，判断是否含有日期，含有则不做处理，不含有则将改行添加到前一行
    '''
    separator = '****'  #分隔符
    Lines = [] #读取存储文本每行
    #读取经过句号分隔的文本
    file_read = open(file_path,'r',encoding='UTF-8')
    line = file_read.readline().strip('\n')
    Lines.append(line) 
    # line = file_read.readline().strip('\n')   #添加ID时需要
    # Lines.append(line)
    while line:
        line = file_read.readline().strip('\n')  #去除换行符
        if line == separator:  #判断是否为分隔标记
            Lines.append(line)
            line = file_read.readline().strip('\n')  
            Lines.append(line)
            # line = file_read.readline().strip('\n')  #添加ID时需要
            # Lines.append(line)
            continue
        datePattern = '\d{1,2}月(?:\d{1,2}日)?'
        DateMatch = re.search(datePattern,line)
        if DateMatch==None:
            Lines[-1] = Lines[-1] + ' ' + line
        else: Lines.append(line)
    file_read.close()
    return Lines
```

```python
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
    separator = '****'  #分隔符
    DLines = []
    for line in Lines:
        DatePattern = '(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?[和及、])+(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?)|(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)|(?:(?<!\~)(?:\d{4}年)?\d{1,2}月(?:\d{1,2}日)?(?:上午|中午|晚)?(?!\~))'
        #SeveralDatePattern = '(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?[和及、])+(?:(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?(?:~(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?)?)'
        dateS = len(re.findall(DatePattern,line))
        if dateS == 0:    
            DLines.append(line)
        elif dateS == 1:
            Pattern = '^(?:\d{4}年)?\d{1,2}月\d{1,2}日(?:上午|中午|晚)?'
            if re.search(Pattern,line):  #查看是否在行首出现月日
                DLines.append(line)
            else:
                Match = re.search(DatePattern,line)
                s = Match.start()
                if DLines[-1] == separator: #判断上一句是否为间隔符
                    DLines.append(line[:s])
                else:DLines[-1] = DLines[-1] + ' ' + line[:s]  ##添加到上一句末尾
                DLines.append(line[s:])       
        else:
            DateMatch = re.finditer(DatePattern,line)
            lastdatesep = 0
            Flag = True #文本以日期开头，则Flag为False
            for date in DateMatch:
                datesep = date.start()
                if datesep != 0:
                    if lastdatesep == 0 and Flag:
                        #if checkdata(DLines[-1]):  #检查上一句是否有日期
                        if DLines[-1] == separator: #判断上一句是否为间隔符
                            DLines.append(line[lastdatesep:datesep])
                        else: DLines[-1] = DLines[-1] + ' ' +line[lastdatesep:datesep]  #添加到上一句末尾
                        lastdatesep = datesep
                    else:
                        DLines.append(line[lastdatesep:datesep])
                        lastdatesep = datesep
                        Flag = True
                else:
                    Flag = False
            DLines.append(line[lastdatesep:])
    return DLines
```

```python
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
    Pattern = '^(?:(\d{4})年)?\d{1,2}月\d{1,2}日$'
    Pattern2 = '(?<=\s)(\S+至)$'
    for var in Lines:
        Match = re.search(Pattern,var)
        if Match:
            Match2 = re.search(Pattern2,OLines[-1])
            if Match2:
                text = Match2.group(1)
                line = text + Match.group()
                OLines[-1] = OLines[-1][:Match2.start()]
                OLines.append(line)
        else: OLines.append(var)
    return OLines
```

```python
def SplitByCTNDate(Lines):
    '''
    连续段日期处理(可选)
    将连续日期，例如：
    1月24日~1月26日到老家过年
    分隔为：
    1月24日到老家过年
    1月25日到老家过年
    1月26日到老家过年
    '''
    CDLines = []
    for var in Lines:
        Pattern = '(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日\~(\d{1,2})月(\d{1,2})日'
        Match = re.search(Pattern,var)
        if Match:
            e = Match.end()
            startMonth = int(Match.group(2))
            startDay = int(Match.group(3))
            endMonth = int(Match.group(4))

            endDay = int(Match.group(5))
            year = 1970
            if Match.group(1) != None:
                year = int(Match.group(1))

            text = var[e:]
            begin = datetime.date(year,startMonth,startDay)
            end = datetime.date(year,endMonth,endDay)
            delta = datetime.timedelta(days=1)
            if Match.group(1) != None:
                    while begin<=end:
                        Ctext = begin.strftime("%Y年%m月%d日")+text
                        CDLines.append(Ctext)
                        begin+=delta
            else:
                while begin<=end:
                    Ctext = begin.strftime("%m月%d日")+text
                    CDLines.append(Ctext)
                    begin+=delta
        else:CDLines.append(var)
    return CDLines
```

```python
def ClearLines(Lines):
    '''去除分隔后小于等于2个字符的段落'''
    ClearLines = []
    for line in Lines:
        if len(line) > 2:
            ClearLines.append(line)
    return ClearLines
```

**1. 简单日期分隔：识别每行，判断是否含有日期，含有则不做处理，不含有则将改行添加到前一行**  
**2. 更多日期分隔：识别每行文本，判断每行文本包含日期行程数量，**

* 若为0个，直接添加；

* 若为1个，判断该日期是否在行首出现，
  * 若在，则直接添加，
  * 不在，则通过日期出现的首位置分隔文本，判断前一行是否为间隔符，
    * 不是则将日期前文本添加到前一行，日期后文本成为独立一行
    * 是则将日期前文本独立成一行，日期后文本成为独立一行；

* 若为2个及以上，通过判断每个日期出现位置，分隔成多行。  

**3. 通过前面的处理，会将日期在句子后面格式的文本错误分隔，例如：  
1月26日早上体感不适 居家休息至1月27日  
分隔成  
1月26日早上体感不适 居家休息至\n1月27日  
因此，我们可以通过识别每行，判断是否只含有日期，将其重新修正为：  
1月26日早上体感不适\n居家休息至1月27日**  

**4. (可选)连续段日期处理  
    将连续日期，例如：  
    1月24日~1月26日到老家过年  
    分隔为：  
    1月24日到老家过年  
    1月25日到老家过年  
    1月26日到老家过年**  
**5. 去除分隔后小于等于2个字符的段落**

调用以上函数返回的部分结果为：

```python
****
第15号确诊病例 女 54岁 系7号病例亲属 常住地湖北武汉
1月17日乘GS7859次航班从武汉到海口 15时33分乘火车C7473车厢04 从美兰到万宁 乘出租车到万宁市万城镇山泉海小区住处
1月18日乘朋友私家车到山根镇水央屈村游玩
1月19日9时左右乘4路公交车到华亚欢乐城购物 11时左右乘4路公交车到住处
1月20日到小区门口的“一心堂”、“兴民药品超市”两家药店购物
1月24日到琼海市人民医院发热门诊就诊并隔离
1月25日确诊 目前正到定点医院隔离治疗
****
出院第30号确诊病例万宁 女 20岁 常住地湖北武汉
1月20日一行8人自驾从汉口到海南
1月21日13时30分左右到海口 之后到海口骑楼小吃街用餐 15时左右到万宁市神州半岛君临海小区住处
1月22日~1月24日居家未外出  
1月25日16时左右被120救护车送到万宁市人民医院发热门诊就诊 并住院隔离治疗
1月27日确诊 
2月12日出院
****
第31号确诊病例万宁 女 27岁 常住地浙江杭州
1月20日乘MF8489次航班从杭州到海口 乘火车到万宁市万城镇山泉海小区入住
1月21日~1月23日居家未外出
1月24日到兴隆温泉酒店进行集中医学观察
1月25日到万宁市中医院住院治疗
1月27日确诊 目前正到定点医院隔离治疗
****
第41号确诊病例万宁 男 73岁 常住地湖北武汉
1月20日与家人一起自驾从武汉到海南
1月21日13时30分左右到达海口 并到海口骑楼小吃街用餐 15时左右自驾到万宁市神州半岛君临海小区住处 
1月22日~1月24日居家未外出
1月25日被120救护车送到万宁市人民医院就诊
1月28日确诊 目前正到定点医院隔离治疗
****
第71号确诊病例万宁 女 51岁 常住地湖北武汉
2020年1月18日乘GS7859次航班从武汉到海口 朋友驾车接到万宁石梅湾某小区家中
1月30日曾步行到附近某餐馆用餐
2月1日被120救护车接到万宁市中医院隔离治疗
2月3日确诊 目前正到定点医院隔离治疗
...
```

处理结果的保存有两种格式：text格式和excel格式  
默认格式为text，保存为excel格式需调用以下函数:

```python
def textToexcel(file_path,Save_Path_textToexcel):
    LinesList = []
    Lines = ''
    separator = '****'
    file_read = open(file_path,'r',encoding='UTF-8')
    line = file_read.readline().strip('\n')
    Lines = Lines + line + '。'
    while line:
        line = file_read.readline().strip('\n')
        if line == separator:
            LinesList.append(Lines)
            Lines = ''
            line = file_read.readline().strip('\n')
            Lines = Lines + line + '。'
        else : Lines = Lines + line + '。'
    file_read.close()

    if not os.path.exists(Save_Path_textToexcel):
        book = xl.Workbook()
        sheet = book.active
        titles = []
        titles.append('Track')
        sheet.append(titles)
        book.save(Save_Path_textToexcel)

    wb = xl.load_workbook(Save_Path_textToexcel)
    sheet = wb['Sheet']
    r = 2
    for row in LinesList:
        sheet.cell(r, column=3).value = row
        r = r + 1
    wb.save(Save_Path_textToexcel)
```

更多代码细节请查看源代码文件。
