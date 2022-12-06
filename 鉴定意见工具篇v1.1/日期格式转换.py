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


if __name__ == '__main__':
    chineseStr = "二〇二〇年十一月二十四日"
    date = ChineseToDate(chineseStr)
    print(date)  # 2020-9-24