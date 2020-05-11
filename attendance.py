import csv
import xlrd3 as xlrd
    
def today(file):
    l = []
    with open(file,'rt') as f: 
        cr = csv.reader(f)
        for row in cr:
            l.append(row)
    f.close()
    return l

def peoples(file,column):
    peopleList = []
    wb = xlrd.open_workbook(filename=file)
    sheet1 = wb.sheet_by_index(0)
    for i in range(1,sheet1.nrows):
        # sheet1.row(i)[列数]
        peopleList.append(sheet1.row(i)[int(column)].value)
    wb.release_resources()
    return peopleList

def judge():
    file = input("请输入成员列表文件名（带后缀名，xlsx格式）\n")
    column = input("请输入成员名所在列数\n")
    peopleList = peoples(file,column)
    #需要删除的成员名
    # peopleList.remove("")
    todayList = today(input("请输入考勤列表文件名（带后缀名，csv格式）\n"))
    t = input("请输入筛选时长\n")
    rpeople = peopleList.copy()
    repeat = []
    i=0
    print("\n")
    for rows in todayList:
        if rows:
            for people in peopleList:
                if(people in rows[0] and len(rows)>1):
                    # <数值
                    if(int(rows[2])<int(t)):
                        print("时长异常：",people,rows[2],rows[3])
                    if(people in rpeople):
                        rpeople.remove(people)
                    else:
                        i-=1
                        repeat.append(people)
                    i+=1
    if repeat:
        print("\n多次出现：",repeat,"\n")
    if rpeople:
        print("今日未在线：",rpeople)
    print("今日总人数：",i)

if __name__ == '__main__':
    judge()
temps=input("\n")