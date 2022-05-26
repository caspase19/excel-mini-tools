import xlrd
import xlwt

subject_tuple = ("总分","语文","数学","英语","物理","化学","生物","政治","历史","地理")

#def level(class_num):
#    if class_num in [7,8,9,10,11]:
#       return 6000, 11000
#    elif class_num in [12,13,14,15]:
#        return 11000, 19000
#    elif class_num in [16,17]:
#        return 19000, 21000

def counts(list1, score1, score2):
    a, b = 0, 0
    for i in range(len(list1)):
        if list1[i] <= score1:
            a += 1
        elif list1[i] <= score2:
            b += 1
    return a, b

def generate(name_list, subject_name):
#   adjust this fomula according to the score sheet
    subject_num = 5 + 4 * subject_tuple.index(subject_name)
    if ranks.cell_value(1,subject_num) != "联考排名":
        input("请调整成绩单，使总分的分数位于第E列，按任意键退出")
        exit()
    list_1 = []
    for i in range(nrows):
        if name_list.count(ranks.cell_value(i,0)) == 1:
            list_1.append(ranks.cell_value(i,subject_num))
    list_1 = [i for i in list_1 if i != ""]
    return list_1

book = xlrd.open_workbook("ranks.xls")
ranks = book.sheet_by_index(0)
book = xlrd.open_workbook("records.xls")
records = book.sheet_by_index(0)
results = xlwt.Workbook(encoding = 'utf-8')
results_sheet = results.add_sheet("结果")
nrows = ranks.nrows
for i in range(1,records.ncols):
    name_list = records.col_values(i, start_rowx=4, end_rowx=None)
    name_list = [i for i in name_list if i != ""]
    list_1 = generate(name_list, records.cell_value(1,i))
    result = counts(list_1,records.cell_value(2,i),records.cell_value(3,i))
    results_sheet.write(4,i,result[0])
    results_sheet.write(5,i,result[1])
results.save("基数.xls")


