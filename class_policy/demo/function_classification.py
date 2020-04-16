# -*- coding:utf-8 -*-
# __author__ = 'Luoyi'
# 创建日期：2018-6-13
# 创建目的：对已分行业的政策文本打功能标签（"其他"文本需要人工判断）

import pandas as pd
import xlrd
import xlwt
import jieba
from xlutils.copy import copy
import pymysql


def function_classification_excel():
    stop_word = pd.read_csv(stopwords_list, encoding='utf-8', header=None, sep='tips')  # 以防逗号被误认为分隔符，因此设置不存在的分隔词
    stop_word = [' ', '', '　', ' '] + list(stop_word[0])  # 将空格符添到停用词中
    directory_words = pd.read_csv(directory_list, encoding='utf-8', header=None, sep='tips')
    replace_words = pd.read_csv(replaceword_list, encoding='utf-8', header=None, sep='tips')
    directory_words = list(directory_words[0])  # 读取自定义词典，转化为list
    replace_words = list(replace_words[0])  # 读取替换为空的词，转换为list
    for word in directory_words:
        jieba.add_word(word)  # 将自定义词典加入分词系统
        # 读取政策data
    data_rd = xlrd.open_workbook(policy_results)  # 读取已分行业的文本
    data_wt = copy(data_rd)
    table_wt = data_wt.get_sheet(0)  # 读取第一个sheet（根据需要修改）
    table = data_rd.sheets()[0]  # 读取第一个sheet（根据需要修改）
    rows = table.nrows  # 工作表行数
    for row in range(1, rows):  # 第一行为表头，从第二行开始读取数据
        policy_title = table.cell(row, 0).value
        class_label1 = table.cell(row, 1).value  # 获取行业标签
        class_label2 = table.cell(row, 2).value  # 获取行业标签
        policy_title = policy_title.replace('\n', ' ')
        print(policy_title)
        processed_title = title_processing(policy_title, replace_words, stop_word)  # 获取处理后的政策标题
        function_label1 = get_function_label(policy_title, processed_title, class_label1)  # 获取功能标签
        function_label2 = get_function_label(policy_title, processed_title, class_label2)  # 获取功能标签
        print(str(function_label1) + '&' + str(function_label2))
        table_wt.write(row, 3, function_label1)
        table_wt.write(row, 4, function_label2)
        data_wt.save(policy_results)

    print("功能标签获取结束！")


def function_classification_sql():
    # 获取标题分词、去停用词、去自定义替换词后的标题list
    stop_word = pd.read_csv(stopwords_list, encoding='utf-8', header=None, sep='tips')  # 以防逗号被误认为分隔符，因此设置不存在的分隔词
    stop_word = [' ', '', '　', ' '] + list(stop_word[0])  # 将空格符添到停用词中
    directory_words = pd.read_csv(directory_list, encoding='utf-8', header=None, sep='tips')
    replace_words = pd.read_csv(replaceword_list, encoding='utf-8', header=None, sep='tips')
    directory_words = list(directory_words[0])  # 读取自定义词典，转化为list
    replace_words = list(replace_words[0])  # 读取替换为空的词，转换为list
    for word in directory_words:
        jieba.add_word(word)  # 将自定义词典加入分词系统
    # 读取政策data
    connection = pymysql.connect(host='127.0.0.1', user='root', password='yeyang112114', db='engineering-brain', charset='utf8',
                                 cursorclass=pymysql.cursors.DictCursor)  # 根据数据库更换账户和密码
    cursor = connection.cursor()

    sql = "SELECT title,industry_label1,industry_label2 FROM 国务院_result"  # 国务院为table名
    cursor.execute(sql)
    result = cursor.fetchall()  # 获取标题
    for item in result:  # item为字典格式
        title = item.get('title')  # 获取每一条政策的标题
        class_label1 = item.get('industry_label1')  # 获取第一个行业标签
        class_label2 = item.get('industry_label2')  # 获取第二个行业标签
        processed_title = title_processing(title, replace_words, stop_word)  # 获取处理后的政策标题
        function_label1 = get_function_label(title, processed_title, class_label1)  # 获取功能标签
        function_label2 = get_function_label(title, processed_title, class_label2)  # 获取功能标签
        sql = "UPDATE 国务院_result SET function_label1 = '%s',function_label2 = '%s' WHERE title = '%s'" % (
            function_label1, function_label2, title)  # 将获取的行业标签写入数据库
        cursor = connection.cursor()
        cursor.execute(sql)
        connection.commit()
    print("功能标签获取结束！")


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


def get_function_label(policy_title, processed_title, class_label):
    # 识别技术类文本的关键词，可自行添加（以这些词结尾的标题）
    technique_word1 = ['剂', '器', '仪', '模', '阀', '管', '板', '珠', '泵', '机', '胺', '酚', '炉', '膜', '计', '件', '液']
    technique_word2 = ['设备', '装备', '系统', '工艺', '材料', '涂料', '管道', '装置', '技术', '晶体', '合金', '电源', '芯片', '组件', '器件', '制备', '制品', '产品', '表盘', '电站', '机组', '表盘', '焊丝', '电路', '电池', '模块', '机床', '教具', '仪表']
    function_label = ''
    # 是否为制造类
    if class_label == '化工' or class_label == '建材' or class_label == '有色金属' or class_label == '机械制造' or  '纺织服装' == class_label or class_label == '轻工' or class_label == '钢铁':
        tecq_word_count = 0
        for item in technique_word1:
            if policy_title[-1:] == item:
                tecq_word_count += 1
                function_label = '技术创新'
        if tecq_word_count == 0:
            for items in technique_word2:
                if policy_title[-2:] == items:
                    tecq_word_count += 1
                    function_label = '技术创新'
        if tecq_word_count == 0:
            count_list = []
            for class_list in manufacture_class_label_list:
                class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
                class_text = list(class_text[0])
                count = 0
                for k in processed_title:
                    for item in class_text:
                        if k == item:
                            count += 1
                count_list.append(count)
            strclass_label = ''
            if sum(count_list) > 0:
                class_label = count_list.index(max(count_list))
                if class_label == 0:
                    str_class_label = '产业化应用'
                if class_label == 1:
                    str_class_label = '产业规划'
                if class_label == 2:
                    str_class_label = '建设生产'
                if class_label == 3:
                    str_class_label = '技术创新'
                if class_label == 4:
                    str_class_label = '投融资'
                if class_label == 5:
                    str_class_label = '行业规范'
                if class_label == 6:
                    str_class_label = '运行监管'
                if class_label == 7:
                    str_class_label = '进出口贸易'
                function_label = str_class_label
            else:
                function_label = '其他'
    # 是否为海关外贸类
    elif class_label == '海关外贸':
        count_list = []
        for class_list in trade_class_label_list:
            class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
            class_text = list(class_text[0])
            count = 0
            for k in processed_title:
                for item in class_text:
                    if k == item:
                        count += 1
            count_list.append(count)

        strclass_label = ''
        if sum(count_list) > 0:
            class_label = count_list.index(max(count_list))
            if class_label == 0:
                str_class_label = '海关口岸'
            if class_label == 1:
                str_class_label = '情况统计'
            if class_label == 2:
                str_class_label = '调查立案'
            if class_label == 3:
                str_class_label = '规划管理'
            if class_label == 4:
                str_class_label = '进出口控制'
            function_label = str_class_label
        else:
            function_label = '其他'
    # 是否为水路船舶类
    elif class_label == '水路船舶':
        tecq_word_count = 0
        for item in technique_word1:
            if policy_title[-1:] == item:
                tecq_word_count += 1
                function_label = '技术创新'
        if tecq_word_count == 0:
            for items in technique_word2:
                if policy_title[-2:] == items:
                    tecq_word_count += 1
                    function_label = '技术创新'
        if tecq_word_count == 0:
            count_list = []
            for class_list in ship_class_label_list:
                class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
                class_text = list(class_text[0])
                count = 0
                for k in processed_title:
                    for item in class_text:
                        if k == item:
                            count += 1
                count_list.append(count)

            strclass_label = ''
            if sum(count_list) > 0:
                class_label = count_list.index(max(count_list))
                if class_label == 0:
                    str_class_label = '工程建设'
                if class_label == 1:
                    str_class_label = '技术创新'
                if class_label == 2:
                    str_class_label = '水路运输'
                if class_label == 3:
                    str_class_label = '船舶管理'
                if class_label == 4:
                    str_class_label = '船舶工业'
                function_label = str_class_label
            else:
                function_label = '其他'
    # 是否为公路汽车类
    elif class_label == '汽车公路':
        tecq_word_count = 0
        for item in technique_word1:
            if policy_title[-1:] == item:
                tecq_word_count += 1
                function_label = '技术改造'
        if tecq_word_count == 0:
            for items in technique_word2:
                if policy_title[-2:] == items:
                    tecq_word_count += 1
                    function_label = '技术改造'
        if tecq_word_count == 0:
            count_list = []
            for class_list in car_class_label_list:
                class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
                class_text = list(class_text[0])
                count = 0
                for k in processed_title:
                    for item in class_text:
                        if k == item:
                            count += 1
                count_list.append(count)

            str_class_label = ''
            if sum(count_list) > 0:
                class_label = count_list.index(max(count_list))
                if class_label == 0:
                    str_class_label = '公路建设'
                if class_label == 1:
                    str_class_label = '制造生产'
                if class_label == 2:
                    str_class_label = '技术改造'
                if class_label == 3:
                    str_class_label = '节能推广'
                if class_label == 4:
                    str_class_label = '道路安全'
                if class_label == 5:
                    str_class_label = '行业运行'
                if class_label == 6:
                    str_class_label = '产业规划'
                function_label = str_class_label
            else:
                function_label = '其他'
    # 是否为铁路类
    elif class_label == '铁路':
        count_list = []
        for class_list in train_class_label_list:
            class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
            class_text = list(class_text[0])
            count = 0
            for k in processed_title:
                for item in class_text:
                    if k == item:
                        count += 1
            count_list.append(count)
        str_class_label = ''
        if sum(count_list) > 0:
            class_label = count_list.index(max(count_list))
            if class_label == 0:
                str_class_label = '安全管理'
            if class_label == 1:
                str_class_label = '建设规划'
            if class_label == 2:
                str_class_label = '铁路运输'
            function_label = str_class_label
        else:
            function_label = '其他'
    # 是否为航空航天类
    elif class_label == '航空航天':
        tecq_word_count = 0
        for item in technique_word1:
            if policy_title[-1:] == item:
                tecq_word_count += 1
                function_label = '技术创新'
        if tecq_word_count == 0:
            for items in technique_word2:
                if policy_title[-2:] == items:
                    tecq_word_count += 1
                    function_label = '技术创新'
        if tecq_word_count == 0:
            count_list = []
            for class_list in aviation_class_label_list:
                class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
                class_text = list(class_text[0])
                count = 0
                for k in processed_title:
                    for item in class_text:
                        if k == item:
                            count += 1
                count_list.append(count)
            str_class_label = ''
            if sum(count_list) > 0:
                class_label = count_list.index(max(count_list))
                if class_label == 0:
                    str_class_label = '制造生产'
                if class_label == 1:
                    str_class_label = '技术创新'
                if class_label == 2:
                    str_class_label = '机场建设'
                if class_label == 3:
                    strclass_label = '行业管理'
                if class_label == 4:
                    str_class_label = '航空运输'
                if class_label == 5:
                    str_class_label = '高技术产业化'
                function_label = str_class_label
            else:
                function_label = '其他'
    # 是否为电子软件类
    elif class_label == '电子软件':
        tecq_word_count = 0
        for item in technique_word1:
            if policy_title[-1:] == item:
                tecq_word_count += 1
                function_label = '技术研究'
        if tecq_word_count == 0:
            for items in technique_word2:
                if policy_title[-2:] == items:
                    tecq_word_count += 1
                    function_label = '技术研究'
        if tecq_word_count == 0:
            count_list = []
            for class_list in digitalInfo_class_label_list:
                class_text = pd.read_csv(class_list, encoding='utf-8', header=None, sep='tips')
                class_text = list(class_text[0])
                count = 0
                for k in processed_title:
                    for item in class_text:
                        if k == item:
                            count += 1
                count_list.append(count)
            str_class_label = ''
            if sum(count_list) > 0:
                class_label = count_list.index(max(count_list))
                if class_label == 0:
                    str_class_label = '产业化'
                if class_label == 1:
                    str_class_label = '工程建设'
                if class_label == 2:
                    str_class_label = '技术研究'
                if class_label == 3:
                    str_class_label = '服务保障'
                if class_label == 4:
                    str_class_label = '监管规划'
                if class_label == 5:
                    str_class_label = '运行统计'
                function_label = str_class_label
            else:
                function_label = '其他'
    elif class_label == '':
        function_label = ''
    else:
        function_label = '无功能分类'

    return function_label


if "__name__ = __main()__":
    policy_results = 'policyClassification/industry_classification/policytitle_results.xls'  # 标题分类结果
    stopwords_list = 'policyClassification/stoplist.txt'  # 停用词典，用于去掉无意义的词
    directory_list = 'policyClassification/directoryByMe.txt'  # 自定义词典，用于保留特殊词组
    replaceword_list = 'policyClassification/replaceWord.txt'  # 自定义替换词，用于替换国务院、工信部等部门抬头
    # 制造业关键词库
    manufacture_class_label_list = ['policyClassification/function_classification/制造Keywords/产业化应用.txt',
                                   'policyClassification/function_classification/制造Keywords/产业规划.txt',
                                   'policyClassification/function_classification/制造Keywords/建设生产.txt',
                                   'policyClassification/function_classification/制造Keywords/技术创新.txt',
                                   'policyClassification/function_classification/制造Keywords/投融资.txt',
                                   'policyClassification/function_classification/制造Keywords/行业规范.txt',
                                   'policyClassification/function_classification/制造Keywords/运行监管.txt',
                                   'policyClassification/function_classification/制造Keywords/进出口贸易.txt']
    # 海关外贸关键词库
    trade_class_label_list = ['policyClassification/function_classification/海关外贸Keywords/海关口岸.txt',
                             'policyClassification/function_classification/海关外贸Keywords/情况统计.txt',
                             'policyClassification/function_classification/海关外贸Keywords/调查立案.txt',
                             'policyClassification/function_classification/海关外贸Keywords/规划管理.txt',
                             'policyClassification/function_classification/海关外贸Keywords/进出口控制.txt']
    # 水路船舶关键词库
    ship_class_label_list = ['policyClassification/function_classification/水路船舶Keywords/工程建设.txt',
                            'policyClassification/function_classification/水路船舶Keywords/技术创新.txt',
                            'policyClassification/function_classification/水路船舶Keywords/水路运输.txt',
                            'policyClassification/function_classification/水路船舶Keywords/船舶管理.txt',
                            'policyClassification/function_classification/水路船舶Keywords/船舶工业.txt']
    # 汽车公路关键词库
    car_class_label_list = ['policyClassification/function_classification/汽车公路Keywords/公路建设.txt',
                           'policyClassification/function_classification/汽车公路Keywords/制造生产.txt',
                           'policyClassification/function_classification/汽车公路Keywords/技术改造.txt',
                           'policyClassification/function_classification/汽车公路Keywords/节能推广.txt',
                           'policyClassification/function_classification/汽车公路Keywords/道路安全.txt',
                           'policyClassification/function_classification/汽车公路Keywords/行业运行.txt',
                           'policyClassification/function_classification/汽车公路Keywords/产业规划.txt']
    # 铁路关键词库
    train_class_label_list = ['policyClassification/function_classification/铁路Keywords/安全管理.txt',
                             'policyClassification/function_classification/铁路Keywords/建设规划.txt',
                             'policyClassification/function_classification/铁路Keywords/铁路运输.txt']
    # 航空航天关键词库
    aviation_class_label_list = ['policyClassification/function_classification/航空航天Keywords/制造生产.txt',
                                'policyClassification/function_classification/航空航天Keywords/技术创新.txt',
                                'policyClassification/function_classification/航空航天Keywords/机场建设.txt',
                                'policyClassification/function_classification/航空航天Keywords/行业管理.txt',
                                'policyClassification/function_classification/航空航天Keywords/航空运输.txt',
                                'policyClassification/function_classification/航空航天Keywords/高技术产业化.txt']
    # 电子软件关键词库
    digitalInfo_class_label_list = ['policyClassification/function_classification/电子软件Keywords/产业化.txt',
                                   'policyClassification/function_classification/电子软件Keywords/工程建设.txt',
                                   'policyClassification/function_classification/电子软件Keywords/技术研究.txt',
                                   'policyClassification/function_classification/电子软件Keywords/服务保障.txt',
                                   'policyClassification/function_classification/电子软件Keywords/监管规划.txt',
                                   'policyClassification/function_classification/电子软件Keywords/运行统计.txt']
    # function_classification_excel()
    function_classification_sql()
