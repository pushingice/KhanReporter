# -*- coding: utf-8 *-*
import os.path

def returnStudentMap():
    """
    returns a map of students
    key: student email
    value: [first, last, full]
    """
    studentMap = {}
    studentPath = os.path.join("private","student_uuids.txt")
    file = open(studentPath).readlines()

    for line in file:
        tokens = line.split()
        studentMap[tokens[-1].strip()] = tokens[0:-1]

    return studentMap


if __name__ == "__main__":

    stMap = returnStudentMap()
    for key in stMap:
        print(key + " " + str(len(stMap[key])))
