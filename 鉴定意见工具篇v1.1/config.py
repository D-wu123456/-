def get_normals():
    with open('./配置信息/规范.txt','r',encoding='utf-8') as read:
        normals=read.read().splitlines()
    return normals
def get_Antivirus():
    with open('./配置信息/天擎版本.txt','r',encoding='utf-8') as read:
        anti_version=read.read().splitlines()
    return anti_version