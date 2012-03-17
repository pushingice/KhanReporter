# -*- coding: utf-8 *-*
import os.path

def returnStudentMap():
    """
    returns a map of students
    key: student email
    value: [first, last, full]
    """
    studentMap = {}
    studentPath = os.path.join("private","student_data.csv")
    file = open(studentPath).readlines()

    for line in file[1:]:
        tokens = line.split(',')
        studentMap[tokens[4].strip()] = tokens[0:4]

    return studentMap


if __name__ == "__main__":

    stMap = returnStudentMap()
    for key in stMap:
        print(key + " " + stMap[key][0])
