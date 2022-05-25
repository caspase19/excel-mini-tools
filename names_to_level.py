import xlrd

def counts(list1, score1, score2):
    a, b = 0, 0
    for i in range(len(list1)):
        if list1[i] <= score1:
            a += 1
        elif list1[i] <= score2:
            b += 1
    return a, b

def generate(name_list,subject_num):
    list_1 = []
    for i in range(nrows):
        if name_list.count(ranks.cell_value(i,0)) == 1:
            list_1.append(ranks.cell_value(i,subject_num))
    return list_1


book = xlrd.open_workbook("ranks.xls")
ranks = book.sheet_by_index(0)
book = xlrd.open_workbook("records.xls")
records = book.sheet_by_index(0)
nrows = ranks.nrows
# 调整以下参数，并验证名单是否正确
name_list = records.col_values(1, start_rowx=29, end_rowx=63)
print(name_list)
print(counts(generate(name_list,42),6000,11000))

