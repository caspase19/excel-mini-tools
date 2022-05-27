# 读取成绩单“ranks.xls”，给出某班分属两个给定分段的人数

import xlrd
import xlwt

def level(class_num):
    if class_num in [7,8,9,10,11]:
        return 6000, 11000
    elif class_num in [12,13,14,15]:
        return 11000, 19000
    elif class_num in [16,17]:
        return 19000, 21000

# 生成某班学生成绩组成的列表
def generate(class_num,subject_num):
    list_1 = []
    for i in range(nrows):
        if ranks.cell_value(i,1) == class_num:
            list_1.append(ranks.cell_value(i,subject_num))
    list_1 = [i for i in list_1 if i != ""]
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

book = xlrd.open_workbook("ranks高一下期中.xls")
ranks = book.sheet_by_index(0)
nrows = ranks.nrows
results = xlwt.Workbook(encoding = 'utf-8')
results_sheet = results.add_sheet("期中总分")
results_sheet.write(0,0,"班级")
results_sheet.write(1,0,"高分段")
results_sheet.write(2,0,"中分段")
for i in range(7,18):
    rank_list = generate(i,5)
    print(i, rank_list)
    result = counts(rank_list, level(i)[0], level(i)[1])
    results_sheet.write(0,i-6,i)
    results_sheet.write(1,i-6,result[0])
    results_sheet.write(2,i-6,result[1])
results.save("期中总分.xls")

