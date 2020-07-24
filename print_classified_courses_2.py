import csv

def load_category():

    category = {}

    with open('./data_2.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile)

        for row in csvreader:
            category[row[0]] = sorted(row[1:])

    return category

classified_courses = load_category()
classified_courses["exercise"] = ['GS01' + str(index).zfill(2) \
        for index in range(1, 15)]
classified_courses["music"] = ['GS02' + str(index).zfill(2) \
        for index in range(1, 13)]
classified_courses["colloquium"] = ['GS9331', 'UC9331']

print("{")
for key in list(classified_courses.keys())[:-1]:
    print('\t\'' + key + '\': ' + str(classified_courses[key]) + ', ')

key = list(classified_courses.keys())[-1]
print('\t\'' + key + '\': ' + str(classified_courses[key]))

print('\t}')
