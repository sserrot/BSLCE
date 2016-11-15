import read_master_file
import read_attendance_file

# RETURNS A DICTIONARY OF STUDENTS WHO DID NOT ATTEND AND THEIR EMAIL


def main():
    non_attendant_students = {}
    i = 0
    first_name_index = 0
    email_index = 2

    master = read_master_file.master()
    attendants = read_attendance_file.attended()

    for person in master:
        if person['pm_id'] not in attendants:
            non_attendant_students[person['first_name']] = person['email']

    return non_attendant_students


