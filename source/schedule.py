import math
import os
import pandas
import xlsxwriter
import configparser


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
CONFIG_FILE_NAME = "schedule.ini"
CONFIG_FILE_PATH = f"{CURRENT_DIRECTORY}/{CONFIG_FILE_NAME}"
TEAMS_FILE_NAME = "teams.csv"
TEAMS_FILE_PATH = f"{CURRENT_DIRECTORY}/{TEAMS_FILE_NAME}"
SCHEDULE_FILE_NAME = "schedule.xlsx"

# The keys in the config file
SEASON_TYPE = "season_type"
SLOT_SIZE = "slot_size"
HALLS_COUNT = "halls_count"
HALLS_NAMES = "halls_names"
RANDOM_SEED = "random_seed"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


# The column indices in the CSV file
ACTIVE_COLUMN_INDEX = get_column_index("A")
STUDENT_COLUMN_INDEX = get_column_index("C")
MENTOR_COLUMN_INDEX = get_column_index("H")
COORDINATOR_COLUMN_INDEX = get_column_index("B")
SEASON_TYPE_INDEX = get_column_index("K")


class Slot:
    def __init__(self, start_row: int, start_col: int, number: int, teams: list):
        self.start_row = start_row
        self.start_col = start_col
        self.number = number
        self.teams = teams


class Team:
    def __init__(self, number: int, student_name: str, mentor_name: str, coordinator_name: str):
        self.number = number
        self.student_name = student_name
        self.mentor_name = mentor_name
        self.coordinator_name = coordinator_name

    def __repr__(self):
        return f"({self.number}, {self.student_name}, {self.mentor_name}, {self.coordinator_name})"


def get_config():
    parser = configparser.ConfigParser(allow_no_value=True)
    parser.read(CONFIG_FILE_PATH, encoding="utf-8")
    general_section = parser["General"]
    config = {
        SEASON_TYPE: general_section[SEASON_TYPE],
        SLOT_SIZE: int(general_section[SLOT_SIZE]),
        HALLS_COUNT: int(general_section[HALLS_COUNT]),
        HALLS_NAMES: [name.strip() for name in general_section[HALLS_NAMES].split(",")],
        RANDOM_SEED: int(general_section[RANDOM_SEED])
    }

    return config


def get_column_name(csv_data, column_index: int):
    column_name = csv_data.columns[column_index]
    return column_name


def get_column_data(csv_data, column_index: int):
    column_name = get_column_name(csv_data, column_index)
    column_data = csv_data[column_name].tolist()
    return column_data


def get_teams(csv_data):
    config = get_config()
    active = get_column_data(csv_data, ACTIVE_COLUMN_INDEX)
    students = get_column_data(csv_data, STUDENT_COLUMN_INDEX)
    mentors = get_column_data(csv_data, MENTOR_COLUMN_INDEX)
    coordinators = get_column_data(csv_data, COORDINATOR_COLUMN_INDEX)
    season_types = get_column_data(csv_data, SEASON_TYPE_INDEX)

    teams = list()
    team_number = 0
    for i in range(len(active)):
        if active[i] == "Активен" and season_types[i] == config[SEASON_TYPE]:
            teams.append(Team(team_number, students[i], mentors[i], coordinators[i]))

    return teams


def create_slots(teams: list):
    # time = pandas.to_datetime("11:00:00")  # str(time.time()
    # minutes_to_add = pandas.Timedelta(minutes=5)
    # time = time + minutes_to_add
    teams_by_coordinator = dict()
    for team in teams:
        coordinator_name = team.coordinator_name.lower()
        if coordinator_name not in teams_by_coordinator:
            teams_by_coordinator[coordinator_name] = list()
        teams_by_coordinator[coordinator_name].append(team)

    print(teams_by_coordinator.keys())


def populate_sheet(start_row, start_col, worksheet, teams):
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

    create_slots(teams)

    # workbook = xlsxwriter.Workbook(SCHEDULE_FILE_NAME)
    # worksheet = workbook.add_worksheet("Schedule")

    # create_slot(0, 0, worksheet, teams[0:10])
    # create_slot(12, 0, worksheet, teams[11:21])
    # create_slot(24, 0, worksheet, teams[22:32])
    # create_slot(36, 0, worksheet, teams[33:43])
    # create_slot(48, 0, worksheet, teams[44:54])

    # worksheet.autofit()
    # workbook.close()


if __name__ == "__main__":
    create_schedule()
