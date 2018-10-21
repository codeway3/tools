#!/usr/bin/env python
# coding: utf-8
# dev-environment: python3.5
# required module: docx, openpyxl
# PS. 部分调试语句没有删除
"""
Usage:
    word_num_counter [-s]

Options:
    -h --help       Show this screen.
    -s              Use cache to run.
"""
import os
import re
import string
import pickle
import docx
from docx import Document
from docopt import docopt
from openpyxl import Workbook
import jieba
import jieba.analyse as ana
import pandas as pd


SOURCEPATH = './tmp/招股说明书/'  # 文档目录
MIDDLEPATH = './tmp/storage.pickle'  # 临时存储位置（可忽略
FINPATH = './tmp/风险段落字数统计.xlsx'  # 字数Excel生成位置
SIPPATH = './tmp/风险段落词数统计.xlsx'  # 词数Excel生成位置
GENPATH = './tmp/风险段落/'  # Word生成位置
STOPWORDS = './tmp/stopword.txt'  # 停用词

dict1 = './tmp/财经金融词汇大全.txt'
dict2 = './tmp/经济 财经 金融 证券 货币 商品 市场 外汇.txt'
jieba.load_userdict(dict1)
jieba.load_userdict(dict2)

# 匹配用的模式串
pattern1_1_in = re.compile(r'^.*重大事项提示.*$')
pattern1_2_in = re.compile(r'^[一二三四五六七八九十]+、.*风险.*')
pattern1_sp = re.compile(r'.*(风险|影响).*')
pattern2_in = re.compile(r'^((第[一二三四五六七八九十]+节)|(第[一二三四五六七八九十]+章))风险因素(及对策)*$')
pattern3_1_in = re.compile(r'^((第[一二三四五六七八九十]+节)|(第[一二三四五六七八九十]+章))管理层讨论与分析$')
pattern3_2_in = re.compile(r'^[一二三四五六七八九十]+、.*(盈利能力.*未来|未来.*盈利能力|盈利能力的|的盈利能力|盈利.*前景|未来趋势).*')

pattern1_1_out = re.compile(r'^(\s*)目(\s*)录(\s*)$')
pattern1_2_out = re.compile(r'^[一二三四五六七八九十]+、.*')
pattern2_out = re.compile(r'^(([一二三四五六七八九十]+、)|(第[一二三四五六七八九十]+节)|(第[一二三四五六七八九十]+章))(发行人|发行人的|公司|本公司)基本情况(\s*)$')
pattern3_1_out = re.compile(r'^((第[一二三四五六七八九十]+节)|(第[一二三四五六七八九十]+章))业务发展.*$')
pattern3_2_out = re.compile(r'^[一二三四五六七八九十]+、.*')


# docx 转换为 list
def read_docx(name):
    file = docx.Document(name)
    para_data = list()
    for paragraph in file.paragraphs:
        para_text = paragraph.text
        para_data.append(para_text)
    return para_data


# 文档字数统计，对于字符串列表和单独字符串有不同的处理逻辑
def word_count(d):
    cnt = 0
    if isinstance(d, list):
        for s in d:
            cnt += str_count(s)
    elif isinstance(d, str):
        cnt += str_count(d)
    return cnt


# 字符串字数统计
def str_count(s):
    count_en = count_dg = count_sp = count_zh = count_pu = 0  # 统一将0赋值给这5个变量
    # s_len = len(s)
    for c in s:
        if c in string.ascii_letters:
            count_en += 1
        elif c.isdigit():
            count_dg += 1
        elif c.isspace():
            count_sp += 1
        elif c.isalpha():
            count_zh += 1
        else:
            count_pu += 1
    return count_zh
    # total_chars = count_zh + count_en + count_sp + count_dg + count_pu
    # if total_chars == s_len:
    #     return('总字数：{0},中文字数：{1},英文字数：{2},空格：{3},数字数：{4},标点符号：{5}'.format(s_len, count_zh, count_en, count_sp, count_dg, count_pu))


def getFileName(path):
    filename = []
    f_list = os.listdir(path)
    for i in f_list:
        if os.path.splitext(i)[1] == '.docx':
            filename.append(i)
    return filename


# 暂存字典中的数据（可忽略
def transfer_word(srcdir, tardir):
    ans = list()
    for filename in getFileName(srcdir):
        tmp = dict()
        tmp['name'] = filename
        tmp['text'] = read_docx(srcdir+filename)
        ans.append(tmp)
    file = open(tardir, 'wb')
    pickle.dump(ans, file)
    file.close()
    return ans


# 反序列化暂存数据（可忽略
def deserialize_dict(dir):
    with open(dir, 'rb') as file:
        tmp_dict = pickle.load(file)
    return tmp_dict


#  处理分词词数统计
def words_count(dat):
    df_list = []
    ans = pd.DataFrame([], columns=['word'])
    ana.set_stop_words(STOPWORDS)
    for para in dat:
        wordlist = ana.extract_tags(para, topK=200)
        df = pd.DataFrame(wordlist, columns=['word'])
        df_list.append(df)
    ans = pd.concat(df_list)
    result = ans.shape[0]
    return result
    # 需要控制输出行数时 改这里 改成 return result.head(xx) xx为需要保留的行数


# 进行字数统计，并输出到表格中
def generate_xlsx(src_dict):
    rows = list()
    crows = list()
    failed_rows = list()
    for dt in src_dict:
        file_name = dt['name']
        row = list()
        crow = list()
        row.append(file_name)
        crow.append(file_name)
        print(file_name)

        all_num = word_count(dt['text'])
        row.append(all_num)
        crow.append(words_count(dt['text']))

        warning_num = [0, 0, 0]
        words_num = [0, 0, 0]
        sumnum = 0
        flag1 = False
        flag3 = False
        flag = [False, False, False]

        document = Document()
        paras_to_word = list()
        for para_text in dt['text']:
            para_text = ''.join(para_text.split())
            if re.match(pattern1_1_out, para_text) and flag1:
                flag1 = False
                flag[0] = False
                print(para_text)
            elif re.match(pattern3_1_out, para_text) and flag3:
                flag3 = False
                flag[2] = False
                print(para_text)
            elif re.match(pattern1_2_out, para_text) and flag1 and flag[0]:
                flag[0] = False
                print(para_text)
            elif re.match(pattern2_out, para_text) and flag[1]:
                flag[1] = False
                print(para_text)
            elif re.match(pattern3_2_out, para_text) and flag3 and flag[2]:
                flag[2] = False
                print(para_text)

            if flag[0]:
                warning_num[0] += word_count(para_text)
                words_num[0] += words_count([para_text])
                paras_to_word.append(para_text)
            elif flag[1]:
                warning_num[1] += word_count(para_text)
                words_num[1] += words_count([para_text])
                paras_to_word.append(para_text)
            elif flag[2]:
                warning_num[2] += word_count(para_text)
                words_num[2] += words_count([para_text])
                paras_to_word.append(para_text)
            elif flag1 and re.match(pattern1_sp, para_text):
                warning_num[0] += word_count(para_text)
                words_num[0] += words_count([para_text])
                paras_to_word.append(para_text)

            if flag[0] or flag[1] or flag[2] or (flag1 and re.match(pattern1_sp, para_text)):
                sumnum += word_count(para_text)

            if re.match(pattern1_2_in, para_text) and flag1:
                flag[0] = True
                print(para_text)
            elif re.match(pattern2_in, para_text):
                flag1 = False
                flag[1] = True
                print(para_text)
            elif re.match(pattern3_2_in, para_text) and flag3:
                flag[2] = True
                print(para_text)
            elif re.match(pattern1_1_in, para_text):
                flag1 = True
                print(para_text)
            elif re.match(pattern3_1_in, para_text):
                flag3 = True
                print(para_text)
        row.extend(warning_num)
        row.append(sum(warning_num))
        crow.extend(words_num)
        crow.append(sum(words_num))
        if sumnum/all_num > 0.8 or warning_num[1] == 0 or warning_num[2] == 0:
            failed_rows.append(row)
        else:
            rows.append(row)
            crows.append(crow)
            # 生成.docx
            for para in paras_to_word:
                if para:
                    document.add_paragraph(para)
            document.save(GENPATH+file_name)
        print()

    wb = Workbook()
    sheet = wb.active
    sheet.title = '主要风险段落字数统计'
    sheet.append(['文档名字', '总字数', '重大事项提示-风险', '风险因素', '管理层讨论与分析-未来盈利能力', '风险段落总字数'])
    for row in rows:
        sheet.append(row)
    wb.create_sheet('处理失败文档', index=1)
    sheet2 = wb['处理失败文档']
    for row in failed_rows:
        sheet2.append(row)
    wb.save(FINPATH)

    wb = Workbook()
    sheet = wb.active
    sheet.title = '主要风险段落词数统计'
    sheet.append(['文档名字', '总词数', '重大事项提示-风险', '风险因素', '管理层讨论与分析-未来盈利能力', '风险段落总词数'])
    for crow in crows:
        sheet.append(crow)
    wb.save(SIPPATH)


if __name__ == '__main__':
    args = docopt(__doc__)
    if args['-s']:
        fd = deserialize_dict(MIDDLEPATH)
    else:
        fd = transfer_word(SOURCEPATH, MIDDLEPATH)
    generate_xlsx(fd)
    print('处理完成')
