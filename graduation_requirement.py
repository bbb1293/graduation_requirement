#@title
#-*- coding:utf-8 -*-

import sys, io
import openpyxl, csv
from termcolor import colored

def get_my_courses():
    ws = openpyxl.load_workbook('Completed course grade.xlsx').active

    index = 6
    ret = []

    while True:
        code = ws['B' + str(index)].value
        credit = ws['E' + str(index)].value
        title = ws['D' + str(index)].value

        if code == '[학사]':
            break;

        if code == None or credit == None or title == None:
            index += 1
            continue;

        ret.append((code, int(credit), title))
        index += 1

    ret.sort(key=lambda x: x[1], reverse=True)
    
    return ret

def load_category():

    category = {}

    with open('./data.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)

        for row in csvreader:
            category[row[0]] = row[1:]

    return category


def classify_my_course(my_course_index):

    # my_course = (code, credit, title)
    my_course = my_courses[my_course_index]

    for category in ["core_english1", "core_english2", "core_math1",
            "core_experiment", "freshman_seminar",
            "others1", "others3"]:

        if my_course[0] in classified_courses[category]:
            my_classified_courses_credit[category] += my_course[1]
            my_classified_courses[category].append(my_course)

            return True

    for category in ["core_writing", "core_math2", "core_science"]:

        if my_course[0] in classified_courses[category]:

            if my_course[1] + my_classified_courses_credit[category] >\
                    classified_courses_credit[category]:
                my_classified_courses_credit["others3"] += my_course[1]
                my_classified_courses["others3"].append(my_course)

            else:
                my_classified_courses_credit[category] += my_course[1]
                my_classified_courses[category].append(my_course)

            return True

    for category in ["HUS", "PPE"]:

        if my_course[0] in classified_courses[category]:

            if my_course[1] + my_classified_courses_credit[category] >\
                    classified_courses_credit[category]:

                if my_course[1] +\
                        my_classified_courses_credit["other_humanity"] >\
                        classified_courses_credit["other_humanity"]:
                    my_classified_courses_credit["others2"] = \
                            my_classified_courses_credit["others2"] +\
                            my_course[1]
                    my_classified_courses["others2"].append(my_course)

                else:
                    my_classified_courses_credit["other_humanity"] +=\
                            my_course[1]
                    my_classified_courses["other_humanity"].append(my_course)

            else:
                my_classified_courses_credit[category] += my_course[1]
                my_classified_courses[category].append(my_course)

            return True

    if my_course[0] in classified_courses["other_humanity"]:

        if my_course[1] +\
                my_classified_courses_credit["other_humanity"] >\
                classified_courses_credit["other_humanity"]:
            my_classified_courses_credit["others2"] += my_course[1]
            my_classified_courses["others2"].append(my_course)

        else:
            my_classified_courses_credit["other_humanity"] +=\
                    my_course[1]
            my_classified_courses["other_humanity"].append(my_course)

        return True

    # Classify my major courses
    for classified_category, category in zip(major[my_major],
            ["major_core", "major_elective"]):

        if my_course[0] in classified_courses[classified_category]:
            my_classified_courses_credit[category] += my_course[1]
            my_classified_courses[category].append(my_course)

            return True

    # Classify other major courses
    for classified_category in [category for sublist in major[:my_major] +\
            major[my_major + 1:] for category in sublist]:

        if my_course[0] in classified_courses[classified_category]:
            my_classified_courses_credit["others3"] += my_course[1]
            my_classified_courses["others3"].append(my_course)

            return True

    # Classify research courses
    for code in classified_courses["research"]:

        if my_course[0][2:] == code:
            my_classified_courses_credit["research"] += my_course[1]
            my_classified_courses["research"].append(my_course)

            return True

    for category in ["music", "exercise", "colloquium"]:

        if my_course[0] in classified_courses[category]:
            my_classified_courses_credit[category] += 1
            my_classified_courses[category].append(my_course)

            return True

    return False

def sum_credits():

    ret = 0

    for category in ["core_english1", "core_english2", "core_writing",
            "HUS", "PPE", "other_humanity",
            "core_math1", "core_math2",
            "core_science", "core_experiment",
            "freshman_seminar", "research",
            "others2"]:
        ret += min(my_classified_courses_credit[category],
                classified_courses_credit[category])

    ret += min(my_classified_courses_credit["major_core"] +\
            my_classified_courses_credit["major_elective"],
            classified_courses_credit["major"])
    ret += (my_classified_courses_credit["others1"] +\
            my_classified_courses_credit["others3"])

    return ret


def print_course(course):
    print('{:<7} {:<7d} {:}'.format(course[0], course[1], course[2]))

def print_courses_by_class(mclass, length = 70):
    print(colored('\n' + courses_name[mclass].center(length) + '\n',
        attrs = ['bold']))
    print('-' * 75)
    print('{:<7} {:<7} {:}'.format("course", "credit", "title"))

    for course in my_classified_courses[mclass]:
        print_course(course)

    print('-' * 75)

    if type(classified_courses_credit[mclass]) is not int or \
            classified_courses_credit[mclass] >\
            my_classified_courses_credit[mclass]:
        print('{:7} {:<7} {:}'.format(" ",
            str(my_classified_courses_credit[mclass]) + '/' + \
                    str(classified_courses_credit[mclass]),
                    courses_text[mclass]))
    else:
        print(colored('{:7} {:<7} {:}'.format(" ",
            str(my_classified_courses_credit[mclass]) + '/' + \
                    str(classified_courses_credit[mclass]),
                    courses_text[mclass]), "red", attrs = ['bold']))

    print('-' * 75 + '\n')

def print_courses_by_subclass(subclass):
    print('- ' + courses_name[subclass] + '\n' + '-' * 75)
    print('{:<7} {:<7} {:}'.format("course", "credit", "title"))

    for course in my_classified_courses[subclass]:
        print_course(course)

    print('-' * 75)

    if subclass == "core_experiment":
        output_string = '{:7} {:<7} {:}'.format(" ",
                str(my_classified_courses_credit[subclass]) + '/' + \
                        str(classified_courses_credit[subclass]),
                        courses_text[subclass])

        if classified_courses_credit[subclass] < 2:
            print(output_string + '\n' + '-' * 75 + '\n')
            return
        else:
            print(colored(output_string, "red", attrs = ['bold']) +\
                    '\n' + '-' * 75 + '\n')
            return
            
    output_string = '{:7} {:<7} {:}'.format(" ",
            str(my_classified_courses_credit[subclass]) + '/' + \
                    str(classified_courses_credit[subclass]),
                    courses_text[subclass])
    if type(classified_courses_credit[subclass]) is not int or \
            classified_courses_credit[subclass] >\
            my_classified_courses_credit[subclass]:
        print(output_string)
    else:
        print(colored(output_string, "red", attrs = ['bold']))

    print('-' * 75 + '\n')

def print_major_courses():

    print(colored('\n' + "전공학점".center(70) + '\n',
        attrs = ['bold']))

    print("- Major Core")
    print('-' * 75)
    print('{:<7} {:<7} {:}'.format("course", "credit", "title"))
    for course in my_classified_courses["major_core"]:
        print_course(course)

    print('-' * 75)
    print('{:7} {:<7}'.format(" ",
        str(my_classified_courses_credit["major_core"])))
    print('-' * 75)

    print("\n- Major elective")
    print('-' * 75)
    print('{:<7} {:<7} {:}'.format("course", "credit", "title"))
    for course in my_classified_courses["major_elective"]:
        print_course(course)

    print('-' * 75)
    print('{:7} {:<7}'.format(" ",
        str(my_classified_courses_credit["major_elective"])))
    print('-' * 75)

    major_credit = my_classified_courses_credit["major_elective"] +\
            my_classified_courses_credit["major_core"]
    
    if major_credit < 30:
        print('{:7} {:<7} {:}'.format(" ",
            str(major_credit) + '/' + \
                    str(classified_courses_credit["major"]),
                    "Mandatory over 30"))
    else:
        print(colored('{:7} {:<7} {:}'.format(" ",
            str(major_credit) + '/' + \
                    str(classified_courses_credit["major"]),
                    "Mandatory over 30"), "red", attrs = ['bold']))

    print('-' * 75 + '\n')

"""
classified_course = [core_english1, core_english2, core_writing,
        HUS, PPE, other_humanity,
        core_math1, core_math2, core_science, core_experiment,
        freshman_seminar,
        physics_core, physics_elective,
        chemical_core, chemical_elective,
        biology_core, biology_elective,
        eecs_core, eecs_elective,
        mechanics_core, mechanics_elective,
        environment_core, environment_elective,
        research,
        others1, others3,
        music, exercise, colloquium]
"""
classified_courses = load_category()
classified_courses["exercise"] = ['GS01' + str(index).zfill(2) \
        for index in range(1, 15)]
classified_courses["music"] = ['GS02' + str(index).zfill(2) \
        for index in range(1, 13)]
classified_courses["colloquium"] = ['GS9331', 'UC9331']
classified_courses_credit = {
        "core_english1": 2, "core_english2": 2, "core_writing": 3,
        "HUS": 6, "PPE": 6, "other_humanity": 12,
        "core_math1": 3, "core_math2": 3,
        "core_science": 9, "core_experiment": 3,
        "freshman_seminar": 1,
        "major": 36,
        "research": 6,
        "others1": '-', "others2": 12, "others3": '-',
        "music": 4, "exercise": 4, "colloquium": 2,
        "nonclassified_courses": '-'
        }
courses_text = {
        **dict.fromkeys(
            ["core_english1", "core_english2", "core_writing",
            "HUS", "PPE", "other_humanity",
            "core_math1", "core_math2", "core_science",
            "freshman_seminar",
            "research",
            "music", "exercise", "colloquium"],
            "Mandatory"),
        **dict.fromkeys(
            ["others1", "others2", "others3"], "Optional"),
        "core_experiment": "Mandatory over 2",
        "nonclassified_courses": ""
        }
courses_name = {
        "core_english1": "Core English 1",
        "core_english2": "Core English 2",
        "core_writing": "Core Writing",
        "HUS": "HUS", "PPE": "PPE",
        "other_humanity": "Other Humanity",
        "core_math1": "Core Mathematics 1",
        "core_math2": "Core Mathematics 2",
        "core_science": "Core Science",
        "core_experiment": "Core Experiment",
        "freshman_seminar": "신입생 세미나",
        "research": "학사논문연구",
        "others1": "자유선택 - 대학 공통과목", 
        "others2": "자유선택 - 인문사회",
        "others3": "자유선택 - 언어선택/소프트웨어, \
기초과학선택, 기초전공, 타 전공",
        "music": "예능실기", "exercise": "체육실기",
        "colloquium": "GIST대학 콜로퀴움",
        "nonclassified_courses": "Nonclassified Courses\
 (학사편람에 기재되지 않은 과목)"
        }
major = [["physics_core", "physics_elective"],
        ["chemical_core", "chemical_elective"],
        ["biology_core", "biology_elective"],
        ["eecs_core", "eecs_elective"],
        ["mechanics_core", "mechanics_elective"],
        ["material_core", "material_elective"],
        ["environment_core", "environment_elective"]]

"""
my_classified_course = [core_english1, core_english2, core_writing,
        HUS, PPE, other_humanity,
        core_math1, core_math2, core_science, core_experiment,
        freshman_seminar,
        major_core, major_elective,
        research,
        others1, other2, others3,
        music, exercise, colloquium, nonclassified_courses]
"""
"""
0: physics, 1: chemical, 2: biology, 3: eecs,
4: mechanics, 5: materials, 6: environment
"""

my_major = 3
my_classified_courses = {category: [] for category in [
    "core_english1", "core_english2", "core_writing",
    "HUS", "PPE", "other_humanity",
    "core_math1", "core_math2", "core_science", "core_experiment",
    "freshman_seminar",
    "major_core", "major_elective",
    "research",
    "others1", "others2", "others3",
    "music", "exercise", "colloquium", "nonclassified_courses"]}
my_classified_courses_credit = {
        **dict.fromkeys(
            ["core_english1", "core_english2", "core_writing",
            "HUS", "PPE", "other_humanity",
            "core_math1", "core_math2", "core_science", "core_experiment",
            "freshman_seminar",
            "major_core", "major_elective",
            "research",
            "others1", "others2", "others3",
            "music", "exercise", "colloquium", "nonclassified_courses"], 0)}

# my_courses = list of (code, credit, title)
my_courses = get_my_courses()

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding = 'utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding = 'utf-8')

    print('-' * 70)
    print("Enter the number corresponding with your major".center(70))
    print("0: Physics, 1: Chemical, 2: Biology, 3: EECS,".center(70))
    print("4: Mechanics, 5: Materials, 6: environment".center(70))
    print('-' * 70)
    my_major = int(input("\n- "))

    for my_course_index in range(len(my_courses)):

        if not classify_my_course(my_course_index):
            my_classified_courses_credit["nonclassified_courses"] +=\
                    my_courses[my_course_index][1]
            my_classified_courses["nonclassified_courses"].append(
                    my_courses[my_course_index])

    print(colored("\n" + "언어의 기초".center(70) + "\n",
        attrs = ['bold']))
    for subclass in ["core_english1", "core_english2", "core_writing"]:
        print_courses_by_subclass(subclass)

    print(colored("\n" + "인문사회".center(70) + "\n",
        attrs = ['bold']))
    for subclass in ["HUS", "PPE", "other_humanity"]:
        print_courses_by_subclass(subclass)

    print(colored("\n" + "기초과학".center(70) + "\n",
        attrs = ['bold']))
    for subclass in ["core_math1", "core_math2", "core_science", 
                     "core_experiment"]:
        print_courses_by_subclass(subclass)

    print_courses_by_class("freshman_seminar")
    print_major_courses()

    for mclass, length in zip(["research", "others1",
        "others2", "others3"], [70, 65, 69, 50]):
        print_courses_by_class(mclass, length)

    print('\n' + colored(('=' * 31).center(75), "red", attrs = ['bold']))
    print(colored("||      Total Credits      ||".center(75), "red",
        attrs = ['bold']))
    print(colored(("{:}{:^25}{:}".format("||", str(sum_credits()) +\
            "/130", "||")).center(75), "red", attrs = ['bold']))
    print(colored(('=' * 31).center(75), "red", attrs = ['bold']))
    print()

    print_courses_by_class("nonclassified_courses", length = 63)
    for mclass in ["music", "exercise", "colloquium"]:
        print_courses_by_class(mclass)
