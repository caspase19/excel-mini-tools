import xlrd
import xlwt

def subject_index(subject_name):
#    subject_tuple = ("总分","语文","数学","英语","物理","化学","生物","政治","历史","地理")
#    subject_num = 6 + 4 * subject_tuple.index(subject_name)

    subject_tuple = ("化学","政治","地理","生物","总分")
    subject_num = 2 + subject_tuple.index(subject_name)

#    subject_num = 5
    return subject_num

def generate(name_list, subject_num):
    rank_list = []
    for i in range(ranks.nrows):
        if name_list.count(ranks.cell_value(i,0)) == 1:
            rank_list.append(ranks.cell_value(i,subject_num))
    rank_list = [i for i in rank_list if i != ""]
    return rank_list

def counts(rank_list, high_level, medium_level):
    a, b = 0, 0
    for i in range(len(rank_list)):
        if rank_list[i] <= high_level:
            a += 1
        elif rank_list[i] <= medium_level:
            b += 1
    return a, b

book = xlrd.open_workbook("ranks.xls")
ranks = book.sheet_by_index(0)
book = xlrd.open_workbook("records.xls")
records = book.sheet_by_index(0)
results = xlwt.Workbook(encoding = 'utf-8')
results_sheet = results.add_sheet("走班基数")
for i in range(4):
    for j in range(records.ncols):
        results_sheet.write(i,j,records.cell_value(i,j))

for i in range(1,records.ncols):
    name_list = records.col_values(i, start_rowx=4, end_rowx=None)
    name_list = [i for i in name_list if i != ""]
    subject_num = subject_index(records.cell_value(1,i))
    rank_list = generate(name_list,subject_num)
    result = counts(rank_list,records.cell_value(2,i),records.cell_value(3,i))
    results_sheet.write(4,i,result[0])
    results_sheet.write(5,i,result[1])
results.save("走班基数.xls")


