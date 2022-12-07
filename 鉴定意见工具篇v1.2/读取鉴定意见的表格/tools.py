import datetime,time
import config
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

    def case_detec(log_name, case, Identification_process, Analysis_Description, Appraisal_opinions, identy):
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
