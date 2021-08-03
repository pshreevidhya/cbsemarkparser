import xlsxwriter
from xlsxwriter import Workbook
import sys

std = sys.argv[1]
filename = sys.argv[2]
output_filename = filename + '.xlsx'
try:
    from itertools import izip_longest  # added in Py 2.6
except ImportError:
    from itertools import zip_longest as izip_longest  # name change in Py 3.x

try:
    from itertools import accumulate  # added in Py 3.2
except ImportError:
    def accumulate(iterable):
        'Return running totals (simplified version).'
        total = next(iterable)
        yield total
        for value in iterable:
            total += value
            yield total


def make_parser(fieldwidths):
    cuts = tuple(cut for cut in accumulate(abs(fw) for fw in fieldwidths))
    # bool values for padding fields
    pads = tuple(fw < 0 for fw in fieldwidths)
    flds = tuple(izip_longest(pads, (0,)+cuts, cuts))[:-1]  # ignore final one

    def parse(line): return tuple(line[i:j].strip()
                                  for pad, i, j in flds if not pad)
    # optional informational function attributes
    parse.size = sum(abs(fw) for fw in fieldwidths)
    parse.fmtstring = ' '.join('{}{}'.format(abs(fw), 'x' if fw < 0 else 's')
                               for fw in fieldwidths)
    return parse


with open(filename, "r") as myfile:
    lines = [line.rstrip() for line in myfile if (
        len(line.strip()) > 0 and line.strip()[0].isdigit())]

students = []
subject = {
    '12': {
        '301': 'English',
        '041': 'Maths',
        '083': 'Science',
        '055': 'Economics',
        '044': 'Enterpreuner',
        '042': 'Computers',
        '043': 'Cooking',
        '054': 'Erumai Meyching',
        '066': 'Dancing',
        '030': 'Kutthaattam'
    },
    '10': {
        '184': 'English',
        '085': 'Hindi',
        '122': 'Sanskrit',
        '018': 'French',
        '041': 'Maths Standard',
        '241': 'Maths Basic',
        '086': 'Science',
        '087': 'Social Science'
    }
}

fileformat = {
    "12": {
        "studentwidth": (11, 2, 52, 8, 8, 8, 8, 8, 9, 3, 3, 6, 26),
        "markwidth": (5, 3, 5, 3, 5, 3, 5, 3, 5, 3, 5, 3)
    },
    "10": {
        "studentwidth": (11, 2, 51, 7, 7, 7, 7, 22, 4),
        "markwidth": (4, 3, 4, 3, 4, 3, 4, 3, 4, 2)
    }
}
for i in range(0, len(lines), 2):
    # negative widths represent ignored padding fields
    studentwidth = fileformat[std]['studentwidth']
    studentparse = make_parser(studentwidth)
    student_tup = studentparse(lines[i])
    markwidth = fileformat[std]['markwidth']
    markparse = make_parser(markwidth)
    marks_tup = markparse(lines[i+1].strip())
    if(std == '10'):
        gr1 = ""
        gr2 = ""
        gr3 = ""
        result = student_tup[8]
    else:
        gr1 = student_tup[9]
        gr2 = student_tup[10]
        gr3 = student_tup[11]
        result = student_tup[12]

    student = {
        "rollno": student_tup[0],
        "gender": student_tup[1],
        "name": student_tup[2],
        "marks": {
            student_tup[3]: {
                "mark": int(marks_tup[0]),
                "grade": marks_tup[1]
            },
            student_tup[4]: {
                "mark": int(marks_tup[2]),
                "grade": marks_tup[3]
            },
            student_tup[5]: {
                "mark": int(marks_tup[4]),
                "grade": marks_tup[5]
            },
            student_tup[6]: {
                "mark": int(marks_tup[6]),
                "grade": marks_tup[7]
            },
            student_tup[7]: {
                "mark": int(marks_tup[8]),
                "grade": marks_tup[9]
            }
        },
        "gr1": gr1,
        "gr2": gr2,
        "gr3": gr3,
        "result": result
    }
    students.append(student)

with open(filename + ".json", "w") as myfile:
    myfile.write(str(students))

wb = Workbook(output_filename)
result_sheet = wb.add_worksheet(std + 'th Results')

result_sheet.write(0, 0, 'Roll Number')
result_sheet.write(0, 1, 'Name')
result_sheet.write(0, 2, 'Gender')
colNumber = 2
for k, v in subject[std].items():
    colNumber = colNumber + 1
    result_sheet.write(0, colNumber, v)
    colNumber = colNumber + 1
    result_sheet.write(0, colNumber, v + " grade")

if(std == '12'):
    colNumber = colNumber + 1
    result_sheet.write(0, colNumber, "gr1")
    colNumber = colNumber + 1
    result_sheet.write(0, colNumber, "gr2")
    colNumber = colNumber + 1
    result_sheet.write(0, colNumber, "gr3")
colNumber = colNumber + 1
result_sheet.write(0, colNumber, "Result")

rowNumber = 0
for student in students:
    rowNumber = rowNumber + 1
    colNumber = 0
    result_sheet.write(rowNumber, colNumber, student['rollno'])
    colNumber = colNumber + 1
    result_sheet.write(rowNumber, colNumber, student['name'])
    colNumber = colNumber + 1
    result_sheet.write(rowNumber, colNumber, student['gender'])
    for subject_code, subject_name in subject[std].items():
        if subject_code in student['marks']:
            m = student['marks'][subject_code]['mark']
            gr = student['marks'][subject_code]['grade']
        else:
            m = ''
            gr = ''
        colNumber = colNumber + 1
        result_sheet.write(rowNumber, colNumber, m)
        colNumber = colNumber + 1
        result_sheet.write(rowNumber, colNumber, gr)
    if(std == '12'):
        colNumber = colNumber + 1
        result_sheet.write(rowNumber, colNumber, student['gr1'])
        colNumber = colNumber + 1
        result_sheet.write(rowNumber, colNumber, student['gr2'])
        colNumber = colNumber + 1
        result_sheet.write(rowNumber, colNumber, student['gr3'])
    colNumber = colNumber + 1
    result_sheet.write(rowNumber, colNumber, student['result'])

print("Total Studends {}".format(len(students)))
wb.close()
