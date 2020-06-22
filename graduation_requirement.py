#-*- coding:utf-8 -*-

import sys, io
import openpyxl, csv

def get_my_courses():
    ws = openpyxl.load_workbook('./Completed course grade.xlsx').active

    index = 6
    ret = []

    while True:
        code = ws['B' + str(index)].value
        
        if code == None:
            index += 1
            continue;

        if code == '[학사]':
            break;

        ret.append((code, int(ws['E' + str(index)].value),
            ws['D' + str(index)].value))

        index += 1

    return ret

def classify_my_course(my_course_index):

    # my_course = (code, credit, title)
    my_course = my_courses[my_course_index]
    
    for i in range(7):

        if my_course[0] in classified_courses[i]:

            if my_course[1] + my_classified_courses_credit[i] > \
                    classified_courses_credit[i]:
                my_classified_courses_credit[-1] += my_course[1]
                my_classified_courses[-1].append(my_course) 

            else:
                my_classified_courses_credit[i] += my_course[1]
                my_classified_courses[i].append(my_course)

            return True

    for i in range(7, 9):
        
        if my_course[0] in classified_courses[i]:

            if my_course[1] + my_classified_courses_credit[i] > \
                    classified_courses_credit[i]:
                
                if my_course[1] + my_classified_courses_credit[9] > \
                        classified_courses_credit[9]:
                    my_classified_courses_credit[-2] = \
                            min(my_classified_courses_credit[-2] + \
                            my_course[1], classified_courses_credit[-2])
                    my_classified_courses[-2].append(my_course)

                else:
                    my_classified_courses_credit[9] += my_course[1]
                    my_classified_courses[9].append(my_course)

            else:
                my_classified_courses_credit[i] += my_course[1]
                my_classified_courses[i].append(my_course)

            return True

    if my_course[0] in classified_courses[9]:
        
        if my_course[1] + my_classified_courses_credit[9] > \
                    classified_courses_credit[9]:
            my_classified_courses_credit[-2] = \
                    min(my_classified_courses_credit[-2] + \
                    my_course[1], classified_courses_credit[-2])
            my_classified_courses[-2].append(my_course)

        else:
            my_classified_courses_credit[9] += my_course[1]
            my_classified_courses[9].append(my_course)

        return True

    if my_course[0] in classified_courses[10]:
        my_classified_courses_credit[10] += my_course[1]
        my_classified_courses[10].append(my_course)

        return True

    return False

def load_category():
    with open("./data.csv", 'r') as csvfile:
        csvreader = csv.reader(csvfile)

        for row in csvreader:
            classified_courses.append(row)

def print_course(course):
    print('{:<7} {:<7d} {:}'.format(course[0], course[1], course[2]))

def print_courses(courses):
    print('{:<7} {:<7} {:}'.format("course", "credit", "title"))
    for course in courses:
        print_course(course)

def print_courses_by_index(index):
    print("\n" + courses_name[index])
    print('-' * 75)
    print_courses(my_classified_courses[index])
    print('-' * 75)
    print('{:7} {:<7} {:}'.format(" ", 
        str(my_classified_courses_credit[index]) + "/" + \
                str(classified_courses_credit[index]), 
                courses_text[index]))
    print('-' * 75)

"""
classified_course = [0.core_math1,
        1.core_math2, 
        2.core_science, 3.core_experiment, 
        4.core_english1, 5.core_english2,
        6.core_writing,
        7.HUS, 8.PPE, 9.GSC,
        10.freshman_seminar]
"""
classified_courses = [['GS1001', 'GS1011'],
        ['GS1002', 'GS2001', 'GS2002', 'GS2004', 'GS2013'],
        ['GS1101', 'GS1103', 'GS1201', 'GS1203', 'GS1301', 'GS1302', 
            'GS1303', 'GS1401'], ['GS1111', 'GS1211', 'GS1311'],
        ['GS1601', 'GS1603'], ['GS1604', 'GS2652'], 
        ['GS1511', 'GS1512', 'GS1513', 
            'GS1531', 'GS1532', 'GS1533', 'GS1534']]
classified_courses_credit = [3, 3, 9, 3, 2, 2, 3, 6, 6, 12, 1, 12, '-']
courses_text = ['Mandatory', 'Mandatory', 'Mandatory', 
        'Mandatory (2 or 3)', 'Mandatory', 'Mandatory', 'Mandatory',
        'Mandatory', 'Mandatory', 'Mandatory', 'Mandatory',
        'Optional', '']
courses_name = ['Core Mathematics 1', 'Core Mathematics 2',
        'Core Science', 'Core Experiment', 'Core English 1',
        'Core English 2', 'Core Writing', 'HUS', 'PPE',
        'Remaining Core Humanities', 'Freshman Seminar',
        'Others2', 'Others3']
load_category()
classified_courses.append(['GS9301', 'GS1901'])

"""
classified_course = [0.core_math1,
        1.core_math2, 
        2.core_science, 3.core_experiment, 
        4.core_english1, 5.core_english2,
        6.core_writing,
        7.HUS, 8.PPE, 9.GSC, 10.freshman_seminar,
        -2.others2, -1.others3]
"""

my_classified_courses = [[] for i in range(13)]
my_classified_courses_credit = [0 for i in range(13)]
my_nonclassified_courses = []

# my_courses = list of (code, credit, title)
my_courses = get_my_courses()

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(
            sys.stdout.detach(), encoding = 'utf-8')
    sys.stderr = io.TextIOWrapper(
            sys.stderr.detach(), encoding = 'utf-8')

    # data_workspace = openpyxl.load_workbook('./data.xlsx').active
    # print(categorize_course(data_workspace, 'GS', 'A', 100))
    
    for my_course_index in range(len(my_courses)):

        if not classify_my_course(my_course_index):
            my_nonclassified_courses.append(my_courses[my_course_index])

    for index in range(13):
        print_courses_by_index(index)

    print("\nmy nonclassified courses\n")
    print_courses(my_nonclassified_courses)
