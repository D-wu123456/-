from docx import Document

def get_table_data_dict(list_list):
    """
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
def get_tables(filename):
    #获取鉴定意见内所有的的表格
    doc = Document(filename)
    tables=doc.tables
    table_text_list=[]
    for i in range(len(tables)):
        tb=tables[i]
        #获取表格的行
        tb_rows=tb.rows
        #读取每一行内容
        for i in range(len(tb_rows)):
            row_data=[]
            row_cells=tb_rows[i].cells
            #读取每一行单元格内容
            for cell in row_cells:
                #单元格内容
                row_data.append(cell.text)
            new_list=[]
            for i in row_data:
                if i not in new_list:
                    new_list.append(i)
            table_text_list.append(new_list)
    #将表格分开为三个，同时过滤掉文中出现的表格，只留下鉴定材料和最后的鉴定人签字部分
    index_list=[]
    for i in range(len(table_text_list)):
        if len(table_text_list[i]) ==0:
            #表格未填写这里注意可以优化
            print('未发现表格内容')
        else:
            if table_text_list[i][0] == '检材编号' and table_text_list[i][1]=='检材信息':
                index_list.append(i)
            if table_text_list[i][0] == '司法鉴定人：':
                index_list.append(i)
    table_list_jiancai=[]#检材列表
    table_list_jiancaizhaop=[]#检材照片列表
    table_list_jiandingren=[]#鉴定人以及签名列表
    #获取到三张表的内容，转存到列表
    for i in range(index_list[1],index_list[2]):
        if '照片' not in table_text_list[i][1]:#过滤检材照片
            table_list_jiancaizhaop.append(table_text_list[i])
    for i in range(index_list[2],len(table_text_list)):
        table_list_jiandingren.append(table_text_list[i])
    for i in range(index_list[0],len(table_list_jiancaizhaop)):
        table_list_jiancai.append(table_text_list[i])
    

    print(table_list_jiancai)
    print(table_list_jiancaizhaop)
    print(table_list_jiandingren)

    dict_jiancai_dict=get_table_data_dict(table_list_jiancai)
    dict_jiancaizhaop_dict=get_table_data_dict(table_list_jiancaizhaop)


    #检测检材信息和检材照片表的信息使用名词是否标准
    print(dict_jiancai_dict)
    print(dict_jiancaizhaop_dict)

    #比对检材信息和后面的检材照片信息表内的数据是否一致，并且记录不一致的内容
    for key_jiancai,value_jiancai in dict_jiancai_dict.items():
        if dict_jiancaizhaop_dict.get(key_jiancai) == value_jiancai:
            continue
        else:
            for key_key,value_value in value_jiancai.items():
                if dict_jiancaizhaop_dict.get(key_jiancai).get(key_key) == value_value:
                    continue
                else:
                    print('编号为:'+repr(key_jiancai)+'的检材,两个表格内的'+key_key+'不一致,分别为:'+repr(value_value)+'和'+repr(dict_jiancaizhaop_dict.get(key_jiancai).get(key_key)))
    #检材表的字典类型数据和检材照片表的字典类型数据以及鉴定人信息
    return dict_jiancai_dict,dict_jiancaizhaop_dict,table_list_jiandingren


filename='./手机、ipad和电脑.docx'
get_tables(filename)