import datetime,time
import config
from docx import Document
#处理检材表格和检材照片表格内的数据
def get_table_data_dict(list_list):
    """主要功能是处理检材表格和检材照片表格内的数据
    传过来列表类型的列表 返回字典类型的字典，方便存取数据
    :param list_list:
    :return:
    """
    dict_data_dict={}
    for i in range(len(list_list)):
        value_dict={}
        for j in range(i+1,len(list_list)):
            if len(list_list[i]) == 3:
                value_dict.setdefault(list_list[i][1], list_list[i][2])
            if len(list_list[i]) == 5:
                value_dict.setdefault(list_list[i][1], list_list[i][2])
                value_dict.setdefault(list_list[i][3], list_list[i][4])
            if list_list[i][0] == list_list[j][0]:
                if len(list_list[j]) == 3:
                    value_dict.setdefault(list_list[j][1],list_list[j][2])
                if len(list_list[j]) == 5:
                    value_dict.setdefault(list_list[j][1],list_list[j][2])
                    value_dict.setdefault(list_list[j][3], list_list[j][4])
            else:
                i=j
                if len(value_dict) > 1:
                    dict_data_dict.setdefault(list_list[i-1][0],value_dict)
                break
    return dict_data_dict
# 获取鉴定意见内所有的的表格
def get_tables(filename):
    """
    此处默认鉴定意见内的两处检材的表格是复制粘贴过去的
    :param filename:
    :return:
    """
    doc = Document(filename)
    tables = doc.tables
    table_text_list = []
    for i in range(len(tables)):
        tb = tables[i]
        # 获取表格的行
        tb_rows = tb.rows
        # 读取每一行内容
        for i in range(len(tb_rows)):
            row_data = []
            row_cells = tb_rows[i].cells
            # 读取每一行单元格内容
            for cell in row_cells:
                # 单元格内容
                row_data.append(cell.text)
            new_list = []
            for i in row_data:
                if i not in new_list:
                    new_list.append(i)
            table_text_list.append(new_list)
    # print(table_text_list)
    # 将表格分开为三个，同时过滤掉文中出现的表格，只留下鉴定材料和最后的鉴定人签字部分
    #这个列表的作用是记录获取的word文档内表格位置

    index_list = []
    for i in range(len(table_text_list)):
        if len(table_text_list[i]) == 0:
            # 表格未填写这里注意可以优化
            print('未发现表格内容')
        else:
            if table_text_list[i][0] == '检材编号' and table_text_list[i][1] == '检材信息':
                index_list.append(i)
            if table_text_list[i][0] == '司法鉴定人：':
                index_list.append(i)
    table_list_jiancai = []  # 检材列表
    table_list_jiancaizhaop = []  # 检材照片列表
    table_list_jiandingren = []  # 鉴定人以及签名列表
    # 获取到三张表的内容，转存到列表
    for i in range(index_list[1], index_list[2]):
        if '照片' not in table_text_list[i][1]:  # 过滤检材照片
            table_list_jiancaizhaop.append(table_text_list[i])
    for i in range(index_list[2], len(table_text_list)):
        table_list_jiandingren.append(table_text_list[i])
    for i in range(index_list[0], len(table_list_jiancaizhaop)):
        table_list_jiancai.append(table_text_list[i])
    #
    # print(table_list_jiancai)
    # print(table_list_jiancaizhaop)
    # print(table_list_jiandingren)

    dict_jiancai_dict = get_table_data_dict(table_list_jiancai)
    dict_jiancaizhaop_dict = get_table_data_dict(table_list_jiancaizhaop)

    # 检测检材信息和检材照片表的信息使用名词是否标准
    # print(dict_jiancai_dict)
    # print(dict_jiancaizhaop_dict)

    # 比对检材信息和后面的检材照片信息表内除了编号其他位置的数据是否一致，并且记录不一致的内容

    # for key_jiancai, value_jiancai in dict_jiancai_dict.items():
    #     if dict_jiancaizhaop_dict.get(key_jiancai) == value_jiancai:
    #         continue
    #     else:
    #         for key_key, value_value in value_jiancai.items():
    #             if dict_jiancaizhaop_dict.get(key_jiancai).get(key_key) == value_value:
    #                 continue
    #             else:
    #                 print('编号为:' + repr(key_jiancai) + '的检材,两个表格内的' + key_key + '不一致,分别为:' + repr(
    #                     value_value) + '和' + repr(dict_jiancaizhaop_dict.get(key_jiancai).get(key_key)))

    # 检材表的字典类型数据和检材照片表的字典类型数据以及鉴定人信息
    return dict_jiancai_dict, dict_jiancaizhaop_dict, table_list_jiandingren
def get_log_name():
    # 获取当前时间
    now_time = datetime.datetime.now()
    # 格式化时间字符串
    str_time = now_time.strftime("%Y-%m-%d %H:%M:%S")
    now_time = time.time()
    return str_time, str(now_time)
def detec_comma(log_,locate,text):
    if not text.endswith('。'):
        tools.write_log(log_, locate, text, '该句子结尾非句号或者缺少句号')
    if text.endswith('。。'):
        tools.write_log(log_, locate, text, '该句话结尾句号数量不正常')
def remove_blank(line):
    list1=['=此页以下为空=','=此页以下为空白=','=以下此页为空=','=以下此页为空白=']
    for text in list1:
        if text in line:
            return True
    return False
class tools():
    def ChineseToDate(chineseStr):
        strch1 = '0一二三四五六七八九十'
        strch2 = '〇一二三四五六七八九十'
        y, m, d = '', '', ''
        if chineseStr.find('年') > 1:
            y = chineseStr[0:chineseStr.index('年')]
        if chineseStr.find('月') > 1:
            m = chineseStr[chineseStr.index('年') + 1:chineseStr.index('月')]
        if chineseStr.find('日') > 1:
            d = chineseStr[chineseStr.index('月') + 1:chineseStr.index('日')]
        # 年
        if len(y) == 4:
            if y.find('0') > 1:
                y = str(strch1.index(y[0:1])) + str(strch1.index(y[1:2])) + str(strch1.index(y[2:3])) + str(
                    strch1.index(y[3:4]))
            else:
                y = str(strch2.index(y[0:1])) + str(strch2.index(y[1:2])) + str(strch2.index(y[2:3])) + str(
                    strch2.index(y[3:4]))
        else:
            return None
        # 月
        if len(m) == 1:
            m = str(strch1.index(m))
        elif len(m) == 2:
            m = str(strch1.index(m[0:1]))[0:1] + str(strch1.index(m[1:2]))

        # 日
        if len(d) == 1:
            d = str(strch1.index(d))
        elif len(d) == 2:
            if len(str(strch1.index(d[0:1]))) == 1:
                d = str(strch1.index(d[0:1])) + str(strch1.index(d[1:2]))[1:2]
            else:
                d = str(strch1.index(d[0:1]))[0:1] + str(strch1.index(d[1:2]))
        elif len(d) == 3:
            d = str(strch1.index(d[0:1])) + str(strch1.index(d[2:3]))
        # 生成 日期
        if y != '' and m != '' and d != '':
            return y + '年' + m + '月' + d + '日'  # datetime.date(int(y), int(m), int(d))
        elif y != '' and m != '':
            return y + '年' + m +'月' # datetime.date(int(y), int(m))
        elif y != '':
            return y + '年'

    def write_log(log_name, locate, detail, error_):
        #记录日志
        with open(log_name + '.txt', 'a+', encoding='utf-8') as wirte:
            wirte.write('\n'+'——>请确认,在段落 <' + locate + '> 的【' + detail + '】\n是否存在 ' + error_ + '\n')

    def case_detec(log_name, case, Identification_process, Analysis_Description, Appraisal_opinions, identy,dict_jiancai,dict_jiancaizhaop):
        # 检验鉴定过程、分析说明、鉴定意见里面的案件编号是否正确
        # 搜索这个元素内包含2022-的元素,取出来判断是不是和案件编号一致
        """
        涉及时间和日期的判断需要优化
        :param case:
        :param Identification_process:
        :param Analysis_Description:
        :param Appraisal_opinions:
        :param identy:
        :return:
        """
        year = case.split('-')[0]
        number = case.split('-')[1]
        number_ = year + '-'
        Identification_detec = []  # 鉴定过程检查结果
        Analysis_detec = []  # 分析说明检查结果
        Appraisal_detec = []  # 鉴定意见检查结果
        identy_detec = []  # 保存和签名检查结果
        error_ = '案件编号错误'

        #首先判断我们的两个表格的检材编号是否有不一致
        # list_jiancai=[]
        # list_jiancaizhaop=[]
        # for key, value in dict_jiancai.items():
        #     list_jiancai.append(key)
        # for key, value in dict_jiancai.items():
        #     list_jiancaizhaop.append(key)
        # #默认两个数组的长度肯定是一致的

        #先检查两个检材表格的检材编号是否与案件编号一致
        for key,value in dict_jiancai.items():
            if case not in key:
                tools.write_log(log_name, '鉴定材料', key, error_)
        for key,value in dict_jiancaizhaop.items():
            if case not in key:
                tools.write_log(log_name, '检材照片', key, error_)

        for text in Identification_process:
            if number_ in text:
                if ('时间' not in text or '日期' not in text) and (case not in text):
                    Identification_detec.append(text)
        if len(Identification_detec) != 0:
            for detail in Identification_detec:
                tools.write_log(log_name, '鉴定过程', detail, error_)
        for text in Analysis_Description:
            if number_ in text:
                if ('时间' not in text or '日期' not in text) and (case not in text):
                    Analysis_detec.append(text)
        if len(Analysis_detec) != 0:
            for detail in Analysis_detec:
                tools.write_log(log_name, '分析说明', detail, error_)
        for text in Appraisal_opinions:
            if number_ in text:
                if ('时间' not in text or '日期' not in text) and (case not in text):
                    Appraisal_detec.append(text)
        if len(Appraisal_detec) != 0:
            for detail in Appraisal_detec:
                tools.write_log(log_name, '鉴定意见', detail, error_)
        for text in identy:
            if year not in text or number not in text:
                identy_detec.append(text)
        if len(identy_detec) != 0:
            for detail in identy_detec:
                tools.write_log(log_name, '数据保存及数字签名', detail, error_)

    def date_detec(log_, date_evaluat, date_acceptance, final_date, prepare_begin):
        #检查鉴定意见中的受理日期，落款日期和鉴定日期是否不一致
        error_ = '时间不一致'
        if date_acceptance not in date_evaluat or final_date not in date_evaluat:
            tools.write_log(log_, '受理日期', date_acceptance, error_)
            tools.write_log(log_, '鉴定日期', date_evaluat, error_)
            tools.write_log(log_, '落款日期', final_date, error_)
        if date_acceptance not in prepare_begin:
            tools.write_log(log_, '受理日期', date_acceptance, error_)
            tools.write_log(log_, '病毒库更新时间', prepare_begin, error_)

    def basic_case_detec(log_, basic_case):
        detec_comma(log_, '基本案情', basic_case)
        # 检查基本案情是否规范
        if not basic_case.startswith('据委托人介绍') and not basic_case.startswith('据委托方介绍'):
            tools.write_log(log_, '基本案情', basic_case, '缺少重要开头:"据委托人（方）介绍"')
        if '我' in basic_case:
            tools.write_log(log_, '基本案情', basic_case, '出现主观词"我"')
        if not basic_case.endswith('。'):
            tools.write_log(log_, '基本案情', basic_case, '结尾有无"。"')
    def ident_matter_detec(log_,ident_matter):
        detec_comma(log_, '委托鉴定事项', ident_matter[len(ident_matter)-1])
        #检查委托鉴定事项1.是否只有一条，多条是否使用封号隔开，，逗号结尾？
        if len(ident_matter) !=1:
            for i in range(len(ident_matter)-1):
                if not ident_matter[i].endswith('；'):
                    tools.write_log(log_, '委托鉴定事项', ident_matter[i], '结尾是否缺少"；"')
            if not ident_matter[len(ident_matter)-1].endswith('。'):
                tools.write_log(log_, '委托鉴定事项', ident_matter[len(ident_matter)-1], '结尾是否缺少"。"')
        else:
            if not ident_matter[0].endswith('。'):
                tools.write_log(log_, '委托鉴定事项', ident_matter[0], '结尾是否缺少"。"')
    def standard_detec(log_,standard):
        detec_comma(log_, '采用的技术标准、技术规范或者技术方法', standard)
        #先检查里面的"、"与"和"的使用是否正确
        #检查规范是否正确，目前无法实现检验当前鉴定意见的规范使用哪一些
        normals=config.get_normals()
        if not standard.startswith('本次鉴定依据') or not standard.endswith('进行检验。'):
            tools.write_log(log_, '采用的技术标准、技术规范或者技术方法', standard, '缺少固定的开头和结尾:"本次鉴定依据"和"进行检验。"')
        else:
            standard=standard.replace('本次鉴定依据','')
            standard=standard.replace('进行检验。','')
            len1=len(standard.split('、'))
            len2=len(standard.split('、')[len(standard.split('、'))-1].split('和'))
            if len1>=2 and len2==2 :
                standard=standard.replace('和','、')
                standard_list=standard.split('、')
                for stan in standard_list:
                    if stan not in normals:
                        tools.write_log(log_, '采用的技术标准、技术规范或者技术方法', stan, '不在鉴定规范列表内，请检查规范的字符并且对照核验"规范.txt"')
            else:
                tools.write_log(log_, '采用的技术标准、技术规范或者技术方法', standard, '使用分隔符"、"与"和"不恰当，格式请参考"A、B、c和D"')
    def instrument_detec(log_,instrument):
        detec_comma(log_, '检验使用的仪器设备', instrument)
        #检查使用的仪器设备格式分隔符等是否正确
        instruments=instrument.split('鉴定使用软件：')[1]
        len1=len(instruments.split('、'))
        len2=len(instruments.split('、')[len(instruments.split('、'))-1].split('和'))
        if not (len1>=2 and len2==2 or len1==1 and len2==2):#修改：也有可能是两个，软件名规范检测
            tools.write_log(log_, '检验使用的仪器设备', instruments, '使用分隔符"、"与"和"不恰当，格式请参考"A、B、c和D"')
    def prepare_begin_detec(log_,prepare_begin):
        detec_comma(log_, '开始前的准备', prepare_begin)
        #检查所写的天擎软件版本是否正确
        anti_version=config.get_Antivirus()
        if anti_version[0] not in prepare_begin:
            tools.write_log(log_, '开始前的准备', prepare_begin,'杀毒软件版本是否有误（使用杀毒软件非天擎请忽略此条）')
    def identy_detec(log_,case,identy,Data_carriers):
        for text in identy:
            detec_comma(log_, '保存和签名',text)
        #交付盘类型检测
        year = case.split('-')[0]
        number = case.split('-')[1]
        text=identy[0]+identy[1]
        CD_=text.find('光盘')
        UD_=text.find('U盘')
        HD_=text.find('硬盘')
        if CD_ !=-1 and UD_ !=-1 and HD_ !=-1:
            tools.write_log(log_, '保存和签名', text,'使用交付盘类型不统一')
        elif CD_ !=-1 and UD_ !=-1:
            tools.write_log(log_, '保存和签名', text,'使用交付盘类型不统一')
        elif CD_ !=-1 and HD_ !=-1:
            tools.write_log(log_, '保存和签名', text,'使用交付盘类型不统一')
        elif UD_ !=-1 and HD_ !=-1:
            tools.write_log(log_, '保存和签名', text,'使用交付盘类型不统一')
        elif CD_ ==-1 and UD_ ==-1 and HD_ ==-1:
            tools.write_log(log_, '保存和签名', text,'使用交付盘类型正确')
        else:
            if CD_ !=-1:
                carry='光盘'
                en_carry='ECD'
                carry_=year+'第'+number+'号'+en_carry
                if '一次性写入的' not in identy[0]:
                    tools.write_log(log_, '保存和签名', identy[0], '光盘要加上“一次性写入的”')
                for line in identy:
                    if carry_ not in line:
                        tools.write_log(log_, '保存和签名', line,'交付盘编号和交付盘类型不一致')
                if carry not in Data_carriers:
                    tools.write_log(log_, '检出数据存放', Data_carriers,'和数据存放盘类型保持一致')
            if UD_ !=-1:
                carry='U盘'
                en_carry='EUD'
                carry_=year+'第'+number+'号'+en_carry
                if '具有写保护措施的' not in identy[0]:
                    tools.write_log(log_, '保存和签名', identy[0], 'U盘要加上“具有写保护措施的”')
                for line in identy:
                    if carry_ not in line:
                        tools.write_log(log_, '保存和签名', line,'交付盘编号和交付盘类型不一致')
                if carry not in Data_carriers:
                    tools.write_log(log_, '检出数据存放', Data_carriers,'和数据存放盘类型保持一致')
            if HD_ !=-1:
                carry='硬盘'
                en_carry='EHD'
                carry_=year+'第'+number+'号'+en_carry
                for line in identy:
                    if carry_ not in line:
                        tools.write_log(log_, '保存和签名', line,'交付盘编号和交付盘类型不一致')
                if carry not in Data_carriers:
                    tools.write_log(log_, '检出数据存放', Data_carriers,'和数据存放盘类型保持一致')
    def Identy_Analy_Appra_detec(log_,Identification_process,Analysis_Description,Appraisal_opinions):
        #分析鉴定过程、分析说明、鉴定意见的问题
        """
        """
