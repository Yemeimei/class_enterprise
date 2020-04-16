# -*- coding:utf-8 -*-
# __author__ = 'Luoyi'
# 创建日期：2018-6-13
# 创建目的：对政策文本打行业标签（"未有合适分类"的文本需要人工判断）

import xlrd
import xlwt
from xlutils.copy import copy
import pandas as pd
import jieba
import pymysql


# 行业分类EXCEL
def industry_classification_excel():
    stop_word = pd.read_csv(stopwords_list, encoding='utf-8', header=None, sep='tips')  # 以防逗号被误认为分隔符，因此设置不存在的分隔词
    stop_word = [' ', '', '　', ' '] + list(stop_word[0])  # 将空格符添到停用词中
    directory_words = pd.read_csv(directory_list, encoding='utf-8', header=None, sep='tips')
    replace_words = pd.read_csv(replaceword_list, encoding='utf-8', header=None, sep='tips')
    directory_words = list(directory_words[0])  # 读取自定义词典，转化为list
    replace_words = list(replace_words[0])  # 读取替换为空的词，转换为list
    for word in directory_words:
        jieba.add_word(word)  # 将自定义词典加入分词系统
    # 读取政策data
    original_text = get_policy_title()
    data_rd = xlrd.open_workbook(original_text)  # 读取原始文本
    data_wt = copy(data_rd)
    table_wt = data_wt.get_sheet(0)  # 读取第一个sheet（根据需要修改）[0到3]
    table = data_rd.sheets()[0]  # 读取第一个sheet（根据需要修改）[0到3]
    rows = table.nrows  # 工作表行数
    for row in range(1, rows):  # 第一行为表头，从第二行开始读取数据
        policy_title = table.cell(row, 0).value
        policy_title = policy_title.replace('\n', ' ')
        print(policy_title)
        processed_title = title_processing(policy_title, replace_words, stop_word)  # 获取处理后的政策标题
        print(processed_title)
        industry_label1, industry_label2 = get_industry_label(processed_title)  # 获取机器标签
        print(str(industry_label1) + '&' + str(industry_label2))
        table_wt.write(row, 1, industry_label1)
        table_wt.write(row, 2, industry_label2)
        data_wt.save(original_text)

    print("行业标签获取结束！")


# 行业分类Mysql
def industry_classification_sql():
    stop_word = pd.read_csv(stopwords_list, encoding='utf-8', header=None, sep='tips')  # 以防逗号被误认为分隔符，因此设置不存在的分隔词
    stop_word = [' ', '', '　', ' '] + list(stop_word[0])  # 将空格符添到停用词中
    directory_words = pd.read_csv(directory_list, encoding='utf-8', header=None, sep='tips')
    replace_words = pd.read_csv(replaceword_list, encoding='utf-8', header=None, sep='tips')
    directory_words = list(directory_words[0])  # 读取自定义词典，转化为list
    replace_words = list(replace_words[0])  # 读取替换为空的词，转换为list
    for word in directory_words:
        jieba.add_word(word)  # 将自定义词典加入分词系统
    # 读取政策data
    connection = pymysql.connect(host='rm-8vbif49m6k7l651e5fo.mysql.zhangbei.rds.aliyuncs.com', user='root', password='Liqin1988', db='python', charset='utf8',
                                 cursorclass=pymysql.cursors.DictCursor)  # 根据数据库更换账户和密码
    cursor = connection.cursor()
    #
    # sql = "CREATE TABLE topic_government_fagai SELECT * FROM topic_government_fagai"#新建一个表
    # cursor.execute(sql)
    # connection.commit()
    # print("新建原始表结束！")

    sql = "SELECT industry FROM topic_info_touziguanxi_shangshisj"  # 国务院为table名
    cursor.execute(sql)
    result = cursor.fetchall()  # 获取标题
    for item in result:  # item为字典格式
        title = item.get('industry')  # 获取每一条政策的标题
        processed_title = title_processing(title, replace_words, stop_word)  # 获取处理后的政策标题
        industry_label1, industry_label2 = get_industry_label(processed_title)  # 获取机器标签
        sql = "UPDATE topic_info_touziguanxi_shangshisj SET test = '%s',function_classify = '%s' WHERE industry = '%s'" % (
            industry_label1, industry_label2, title)  # 将获取的行业标签写入数据库
        cursor = connection.cursor()
        cursor.execute(sql)
        connection.commit()
    print("行业标签获取结束！")


# 将原始文件提取政策标题，转化为新excel
def get_policy_title():
    # 读取政策data
    data_rd = xlrd.open_workbook(original_text)  # 读取原始文本
    table = data_rd.sheets()[0]  # 读取第一个sheet（根据需要修改）[0到3]
    rows = table.nrows  # 工作表行数
    data_wt = xlwt.Workbook(encoding='utf-8', style_compression=0)
    table_wt = data_wt.add_sheet('政策分类结果')  # 根据需求改sheet名称
    table_wt.write(0, 0, 'policyTitle')
    table_wt.write(0, 1, 'classlabel1')
    table_wt.write(0, 2, 'classlabel2')
    table_wt.write(0, 3, 'functionlabel1')
    table_wt.write(0, 4, 'functionlabel2')
    for row in range(1, rows):
        policy_title = table.cell(row, 2).value  # 政策标题的列数:国务院2,工信部7,发改委2,商务部1
        table_wt.write(row, 0, policy_title)
    data_wt.save(policy_results)
    print("写入政策标题结束！")
    return policy_results


# 获取标题分词、去停用词、去自定义替换词后的标题list
def title_processing(policy_title, replace_words, stop_word):
    title_list = []
    text = policy_title  # 要处理的原始政策标题
    for word in replace_words:
        text = text.replace(word, "")
    text = ' '.join(jieba.cut(text, cut_all=True))  # 全模式jieba分词
    text = text.split(' ')
    for item in text:
        if item not in stop_word:
            title_list.append(item)
    title = title_list
    return title


# 通过机器给政策两个行业标签
def get_industry_label(title):
    count_list = []
    for class_list in classlabel_list:
        class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tipdm')
        class_text = list(class_text[0])  # 获取关键词库中的词列表
        count = 0
        for k in title:
            for line in class_text:
                line_list = (str(line)).split("&", 1)
                word = line_list[0]  # 关键词
                weight = int(line_list[1])  # 关键词权重
                if k == word:
                    count += weight
        count_list.append(count)

    str_class_label = ''
    if sum(count_list) > 0:
        str_sub_class_label = ''
        max_count = max(count_list)
        class_label = count_list.index(max(count_list))
        if class_label == 0:  # A
            str_class_label = '农林牧渔'
        if class_label == 1:  # B
            str_class_label = '石油天然气矿产'
        if class_label == 2:  # C
            str_class_label = '工业其他'
        if class_label == 3:  # C
            str_class_label = '化工'
        if class_label == 4:  # C
            str_class_label = '建材'
        if class_label == 5:  # C
            str_class_label = '有色金属'
        if class_label == 6:  # C
            str_class_label = '机械制造'
        if class_label == 7:  # C
            str_class_label = '纺织服装'
        if class_label == 8:  # C
            str_class_label = '轻工'
        if class_label == 9:  # C
            str_class_label = '钢铁'
        if class_label == 10:  # C
            str_class_label = '食品农副'
        if class_label == 11:  # D
            str_class_label = '科技专利'
        if class_label == 12:  # E
            str_class_label = '海关外贸'
        if class_label == 13:  # F
            str_class_label = '医药卫生'
        if class_label == 14:  # F
            str_class_label = '教育'
        if class_label == 15:  # F
            str_class_label = '居民生活'
        if class_label == 16:  # G
            str_class_label = '航空航天'
        if class_label == 17:  # G
            str_class_label = '铁路'
        if class_label == 18:  # G
            str_class_label = '水路船舶'
        if class_label == 19:  # G
            str_class_label = '汽车公路'
        if class_label == 20:  # G
            str_class_label = '管道仓储联运'
        if class_label == 21:  # H
            str_class_label = '电子软件'
        if class_label == 22:  # I
            str_class_label = '金融财政'
        if class_label == 23:  # J
            str_class_label = '电热燃力'
        if class_label == 24:  # K
            str_class_label = '土地水利气象'
        if class_label == 25:  # L
            str_class_label = '公共管理'
        if class_label == 26:  # M
            str_class_label = '环保节能'
        if class_label == 27:  # N
            str_class_label = '相关宏观政策'
        if class_label == 28:  # O
            str_class_label = '政务国防组织'

        count_list[class_label] = -1
        if (max(count_list)) == max_count:
            sub_class_label = count_list.index(max(count_list))
            if sub_class_label == 0:  # A
                str_sub_class_label = '农林牧渔'
            if sub_class_label == 1:  # B
                str_sub_class_label = '石油天然气矿产'
            if sub_class_label == 2:  # C
                str_sub_class_label = '工业其他'
            if sub_class_label == 3:  # C
                str_sub_class_label = '化工'
            if sub_class_label == 4:  # C
                str_sub_class_label = '建材'
            if sub_class_label == 5:  # C
                str_sub_class_label = '有色金属'
            if sub_class_label == 6:  # C
                str_sub_class_label = '机械制造'
            if sub_class_label == 7:  # C
                str_sub_class_label = '纺织服装'
            if sub_class_label == 8:  # C
                str_sub_class_label = '轻工'
            if sub_class_label == 9:  # C
                str_sub_class_label = '钢铁'
            if sub_class_label == 10:  # C
                str_sub_class_label = '食品农副'
            if sub_class_label == 11:  # D
                str_sub_class_label = '科技专利'
            if sub_class_label == 12:  # E
                str_sub_class_label = '海关外贸'
            if sub_class_label == 13:  # F
                str_sub_class_label = '医药卫生'
            if sub_class_label == 14:  # F
                str_sub_class_label = '教育'
            if sub_class_label == 15:  # F
                str_sub_class_label = '居民生活'
            if sub_class_label == 16:  # G
                str_sub_class_label = '航空航天'
            if sub_class_label == 17:  # G
                str_sub_class_label = '铁路'
            if sub_class_label == 18:  # G
                str_sub_class_label = '水路船舶'
            if sub_class_label == 19:  # G
                str_sub_class_label = '汽车公路'
            if sub_class_label == 20:  # G
                str_sub_class_label = '管道仓储联运'
            if sub_class_label == 21:  # H
                str_sub_class_label = '电子软件'
            if sub_class_label == 22:  # I
                str_sub_class_label = '金融财政'
            if sub_class_label == 23:  # J
                str_sub_class_label = '电热燃力'
            if sub_class_label == 24:  # K
                str_sub_class_label = '土地水利气象'
            if sub_class_label == 25:  # L
                str_sub_class_label = '公共管理'
            if sub_class_label == 26:  # M
                str_sub_class_label = '环保节能'
            if sub_class_label == 27:  # N
                str_sub_class_label = '相关宏观政策'
            if sub_class_label == 28:  # O
                str_sub_class_label = ' 政务国防组织'
        if str_sub_class_label == '':
            class_label1 = str_class_label
            class_label2 = '无'
        else:
            class_label1 = str_class_label
            class_label2 = str_sub_class_label

    else:
        class_label1 = '未有合适分类'
        class_label2 = '未有合适分类'

    return class_label1, class_label2


if "__name__ = __main()__":
    original_text = 'policyClassification/industry_classification/originalData.xlsx'  # 原始Excel文本
    stopwords_list = 'policyClassification/stoplist.txt'  # 停用词典，用于去掉无意义的词
    directory_list = 'policyClassification/directoryByMe.txt'  # 自定义词典，用于保留特殊词组
    replaceword_list = 'policyClassification/replaceWord.txt'  # 自定义替换词，用于替换国务院、工信部等部门抬头
    # 行业关键词库
    classlabel_list = ['policyClassification/industry_classification/industry_keywords/A农林牧渔.txt',
                       'policyClassification/industry_classification/industry_keywords/B石油天然气矿产.txt',
                       'policyClassification/industry_classification/industry_keywords/C工业其他.txt',
                       'policyClassification/industry_classification/industry_keywords/C化工.txt',
                       'policyClassification/industry_classification/industry_keywords/C建材.txt',
                       'policyClassification/industry_classification/industry_keywords/C有色金属.txt',
                       'policyClassification/industry_classification/industry_keywords/C机械制造.txt',
                       'policyClassification/industry_classification/industry_keywords/C纺织服装.txt',
                       'policyClassification/industry_classification/industry_keywords/C轻工.txt',
                       'policyClassification/industry_classification/industry_keywords/C钢铁.txt',
                       'policyClassification/industry_classification/industry_keywords/C食品农副.txt',
                       'policyClassification/industry_classification/industry_keywords/D科技专利.txt',
                       'policyClassification/industry_classification/industry_keywords/E海关外贸.txt',
                       'policyClassification/industry_classification/industry_keywords/F医药卫生.txt',
                       'policyClassification/industry_classification/industry_keywords/F教育.txt',
                       'policyClassification/industry_classification/industry_keywords/F居民生活.txt',
                       'policyClassification/industry_classification/industry_keywords/G航空航天.txt',
                       'policyClassification/industry_classification/industry_keywords/G铁路.txt',
                       'policyClassification/industry_classification/industry_keywords/G水路船舶.txt',
                       'policyClassification/industry_classification/industry_keywords/G汽车公路.txt',
                       'policyClassification/industry_classification/industry_keywords/G管道仓储联运.txt',
                       'policyClassification/industry_classification/industry_keywords/H电子软件.txt',
                       'policyClassification/industry_classification/industry_keywords/I金融财政.txt',
                       'policyClassification/industry_classification/industry_keywords/J电热燃力.txt',
                       'policyClassification/industry_classification/industry_keywords/K土地水利气象.txt',
                       'policyClassification/industry_classification/industry_keywords/L公共管理.txt',
                       'policyClassification/industry_classification/industry_keywords/M环保节能.txt',
                       'policyClassification/industry_classification/industry_keywords/N相关宏观政策.txt',
                       'policyClassification/industry_classification/industry_keywords/O政务国防组织.txt']  # 关键词库
    policy_results = 'policyClassification/industry_classification/policytitle_results.xls'  # 标题分类结果

    # industry_classification_excel()#原始文件为excel
    industry_classification_sql()  # 原始文件为数据库
