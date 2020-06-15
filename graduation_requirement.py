#-*- coding:utf-8 -*-

import openpyxl
import sys
import io

def categorize_course(ws, substr, col, end_row):
    output = set()
    
    for row in range(2, end_row + 1):
        tmp_index = 0
        val = ws[col + str(row)].value

        if val == None:
            break;

        while val.find(substr, tmp_index) != -1:
            tmp_index = val.find(substr, tmp_index) + 1
            output.add(val[tmp_index - 1:tmp_index + 5])

    return output

def my_courses():
    ws = openpyxl.load_workbook('./Completed course grade.xlsx').active

    index = 6
    codes = []
    titles = []
    credits = []

    while True:
        code = ws['B' + str(index)].value
        
        if code == None:
            index += 1
            continue;

        if code == '[학사]':
            break;

        codes.append(code)
        titles.append(ws['D' + str(index)].value)
        credits.append(ws['E' + str(index)].value)

        index += 1

    return [codes, credits, titles]

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(
            sys.stdout.detach(), encoding = 'utf-8')
    sys.stderr = io.TextIOWrapper(
            sys.stderr.detach(), encoding = 'utf-8')

    data_workspace = openpyxl.load_workbook('./data.xlsx').active
    # print(categorize_course(data_workspace, 'GS', 'A', 100))
    
    courses, credits, titles = my_courses()

    print('{:<7} {:<7} {:}'.format('code', 'credit', 'title'))
    for course, credit, title in zip(courses, credits, titles):
        print('{:<7} {:<7} {:}'.format(course, credit, title))
    
