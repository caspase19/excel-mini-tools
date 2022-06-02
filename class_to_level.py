from openpyxl import Workbook
from openpyxl import load_workbook

CLASS_NUM_COL = 3
TOTAL_SCORE_COL = 6
EXAM = "高一上期末"
PARTS = ((7,8,9,10,11),(12,13,14,15),(16,17))

def level(classNum):
    if classNum in PARTS[0]:
        return 6000, 11000
    elif classNum in PARTS[1]:
        return 11000, 19000
    elif classNum in PARTS[2]:
        return 19000, 21000

# Generate a list of students' ranks, remove absent students'
def generate(classNum,subjectNum):
    rankList = []
    for i in ranks.values:
        if i[CLASS_NUM_COL] == classNum:# adjust this type (int/str) 
            rankList.append([i[TOTAL_SCORE_COL], i[subjectNum]])
    rankList = [i for i in rankList if i[0] != "" and i[1] != ""]
    return rankList

def counts(rankList, highLevel, mediumLevel):
    a, b = 0, 0
    for i in rankList:
        if i[0] <= highLevel:
            if i[1] <= highLevel:
                a += 1
        elif i[0] <= mediumLevel:
            if i[1] <= mediumLevel:
                b += 1
    return a, b

def write(numList):
    for i in numList:
        for j in range(9):
            rankList = generate(i,TOTAL_SCORE_COL+4+4*j)
            result = counts(rankList, level(i)[0], level(i)[1])
            results.active.cell(row=a+i-numList[0], column=2+j, value=result[0])
            results.active.cell(row=a+i-numList[0]+len(numList), column=2+j, value=result[1])


book = load_workbook(f"成绩单/{EXAM}.xlsx")
ranks = book.active
results = Workbook()
for i in PARTS:
    a = 11 * PARTS.index(i) + 1
    write(i)
results.save(f"{EXAM}基数.xlsx")

