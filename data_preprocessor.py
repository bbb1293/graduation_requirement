#-*- coding:utf-8 -*-

import sys, io
import openpyxl, csv

def categorize_course(ws, substr, col, end_row = 100):
    output = set()

    for row in range(2, end_row + 1):
        tmp_index = 0
        val = ws[col + str(row)].value

        if val == None:
            break;

        while val.find(substr, tmp_index) != -1:
            tmp_index = val.find(substr, tmp_index) + 1
            output.add(val[tmp_index - 1: tmp_index + 5])

    return list(output)

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(
            sys.stdout.detach(), encoding = 'utf-8')
    sys.stderr = io.TextIOWrapper(
            sys.stderr.detach(), encoding = 'utf-8')

    workspace = openpyxl.load_workbook('./data.xlsx').active

    with open("./data.csv", 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerows([categorize_course(workspace, 'GS', ch) \
                for ch in "ABC"])
