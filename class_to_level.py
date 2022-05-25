# 读取成绩单“ranks.xls”，给出某班分属两个给定分段的人数

import xlrd

# 生成某班学生成绩组成的列表
def generate(class_num,subject_num):
    list_1 = []
    for i in range(nrows):
        if ranks.cell_value(i,3) == class_num:
            list_1.append(ranks.cell_value(i,subject_num))
    return list_1

# 统计列表中分属两个给定分段的人数
def counts(list1, score1, score2):
    a, b = 0, 0
    for i in range(len(list1)):
        if list1[i] <= score1:
            a += 1
        elif list1[i] <= score2:
            b += 1
    return a, b

book = xlrd.open_workbook("ranks.xls")
ranks = book.sheet_by_index(0)
nrows = ranks.nrows
# 调整以下参数,物理为22,地理为42
print(counts(generate(17,22),19000,21000))

