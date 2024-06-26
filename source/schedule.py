import math
import os
import pandas
import xlsxwriter


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
TEAMS_FILE_NAME = "teams.csv"
TEAMS_FILE_PATH = f"{CURRENT_DIRECTORY}/{TEAMS_FILE_NAME}"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


ACTIVE_COLUMN_INDEX = get_column_index("A")
STUDENT_COLUMN_INDEX = get_column_index("B")
MENTOR_COLUMN_INDEX = get_column_index("H")
COORDINATOR_COLUMN_INDEX = get_column_index("D")


def get_column_name(csv_data, column_index: int):
    column_name = csv_data.columns[column_index]
    return column_name


def get_column_data(csv_data, column_index: int):
    column_name = get_column_name(csv_data, column_index)
    column_data = csv_data[column_name].tolist()
    return column_data


def get_teams(csv_data):
    active = get_column_data(csv_data, ACTIVE_COLUMN_INDEX)
    students = get_column_data(csv_data, STUDENT_COLUMN_INDEX)
    mentors = get_column_data(csv_data, MENTOR_COLUMN_INDEX)
    coordinators = get_column_data(csv_data, COORDINATOR_COLUMN_INDEX)

    teams = list()
    team_number = 1
    time = pandas.to_datetime("11:00:00")
    for i in range(len(active)):
        if active[i] == "Активен":
            teams.append((team_number, students[i], mentors[i], coordinators[i], str(time.time())))
            team_number += 1
            minutes_to_add = pandas.Timedelta(minutes=5)
            time = time + minutes_to_add

    return teams


def create_slot(start_row, start_col, worksheet, teams):
    worksheet.write(start_row, start_col, "#")
    worksheet.write(start_row, start_col + 1, "Ученик")
    worksheet.write(start_row, start_col + 2, "Ментор")
    worksheet.write(start_row, start_col + 3, "Координатор")

    teams_count = len(teams)
    for r in range(teams_count):
        row = start_row + r + 1
        team = teams[r]
        tuple_length = len(team)
        for c in range(tuple_length):
            col = start_col + c
            worksheet.write(row, col, team[c])


def create_schedule():
    csv_data = pandas.read_csv(TEAMS_FILE_PATH)
    teams = get_teams(csv_data)

    workbook = xlsxwriter.Workbook("schedule.xlsx")
    worksheet = workbook.add_worksheet("Schedule")

    create_slot(0, 0, worksheet, teams[0:10])
    create_slot(11, 0, worksheet, teams[11:21])
    create_slot(22, 0, worksheet, teams[22:32])
    create_slot(33, 0, worksheet, teams[33:43])
    create_slot(44, 0, worksheet, teams[44:54])

    worksheet.autofit()

    workbook.close()


if __name__ == "__main__":
    create_schedule()
