import math
import os
import pandas
import xlsxwriter
import configparser
import random


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
CONFIG_FILE_NAME = "schedule.ini"
CONFIG_FILE_PATH = f"{CURRENT_DIRECTORY}/{CONFIG_FILE_NAME}"
TEAMS_FILE_NAME = "teams.csv"
TEAMS_FILE_PATH = f"{CURRENT_DIRECTORY}/{TEAMS_FILE_NAME}"
SCHEDULE_FILE_NAME = "schedule_{0}.xlsx"

# Config section
SEASON_SOFIA = "Sofia"
SEASON_ONLINE = "Online"

# Config keys
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
SEASON_TYPE_COLUMN_INDEX = get_column_index("K")


class Slot:
    def __init__(self, hall_name: str, number: int, teams: list):
        self.hall_name = hall_name
        self.number = number
        self.teams = teams

    def __repr__(self):
        return f"({self.hall_name}, {self.number}, {self.teams})"


class Team:
    def __init__(self, number: int, student_name: str, mentor_name: str, coordinator_name: str):
        self.number = number
        self.student_name = student_name
        self.mentor_name = mentor_name
        self.coordinator_name = coordinator_name

    def __repr__(self):
        return f"({self.number}, {self.student_name}, {self.mentor_name}, {self.coordinator_name})"


def get_config(season: str):
    parser = configparser.ConfigParser(allow_no_value=True)
    parser.read(CONFIG_FILE_PATH, encoding="utf-8")
    config_section = parser[season]
    config = {
        SEASON_TYPE: config_section[SEASON_TYPE],
        SLOT_SIZE: int(config_section[SLOT_SIZE]),
        HALLS_COUNT: int(config_section[HALLS_COUNT]),
        HALLS_NAMES: [name.strip() for name in config_section[HALLS_NAMES].split(",")],
        RANDOM_SEED: int(config_section[RANDOM_SEED])
    }

    return config


def get_column_name(csv_data, column_index: int):
    column_name = csv_data.columns[column_index]
    return column_name


def get_column_data(csv_data, column_index: int):
    column_name = get_column_name(csv_data, column_index)
    column_data = csv_data[column_name].tolist()
    return column_data


def get_teams(config, csv_data):
    active = get_column_data(csv_data, ACTIVE_COLUMN_INDEX)
    students = get_column_data(csv_data, STUDENT_COLUMN_INDEX)
    mentors = get_column_data(csv_data, MENTOR_COLUMN_INDEX)
    coordinators = get_column_data(csv_data, COORDINATOR_COLUMN_INDEX)
    season_types = get_column_data(csv_data, SEASON_TYPE_COLUMN_INDEX)

    teams = list()
    team_number = 0
    for i in range(len(active)):
        if active[i] == "Активен" and season_types[i] == config[SEASON_TYPE]:
            teams.append(Team(team_number, students[i], mentors[i], coordinators[i]))

    return teams


def create_slots(config, teams: list):
    teams_by_coordinator = dict()
    for team in teams:
        coordinator_name = team.coordinator_name.lower()
        if coordinator_name not in teams_by_coordinator:
            teams_by_coordinator[coordinator_name] = list()
        teams_by_coordinator[coordinator_name].append(team)

    halls_count = config[HALLS_COUNT]
    halls_names = config[HALLS_NAMES]
    coordinators_names = list(teams_by_coordinator.keys())
    random.shuffle(coordinators_names)
    coordinators_count = len(coordinators_names)
    teams_count = len(teams)
    teams_count_per_hall = math.ceil(teams_count / halls_count)
    teams_by_hall_name = dict()

    last_coordinator_index = -1
    for i_hall in range(0, halls_count):
        teams_in_hall = list()
        for i_coordinator in range(last_coordinator_index + 1, coordinators_count):
            last_coordinator_index = i_coordinator
            coordinator_name = coordinators_names[i_coordinator]
            teams_in_hall += teams_by_coordinator[coordinator_name]
            teams_in_hall_count = len(teams_in_hall)
            if teams_in_hall_count >= teams_count_per_hall or i_coordinator == coordinators_count - 1:
                random.shuffle(teams_in_hall)
                hall_name = halls_names[i_hall]
                teams_by_hall_name[hall_name] = teams_in_hall
                break

    slot_size = config[SLOT_SIZE]
    slots_by_hall_name = dict()
    for hall_name in halls_names:
        slots_by_hall_name[hall_name] = list()
        teams_in_hall = teams_by_hall_name[hall_name]
        teams_in_hall_count = len(teams_in_hall)
        teams_in_slot = list()
        slot_number = 1
        team_number = 1
        for i_team in range(0, teams_in_hall_count):
            team = teams_in_hall[i_team]
            team.number = team_number
            team_number += 1
            teams_in_slot.append(team)
            teams_in_slot_count = len(teams_in_slot)
            is_slot_full = (teams_in_slot_count == slot_size)
            is_last_team = (i_team == teams_in_hall_count - 1)
            if is_slot_full or is_last_team:
                slot = Slot(hall_name, slot_number, teams_in_slot)
                slots_by_hall_name[hall_name].append(slot)
                teams_in_slot = list()
                slot_number += 1

    return slots_by_hall_name


def populate_sheet(config, worksheet, slots_by_hall_name: dict):
    # time = pandas.to_datetime("11:00:00")  # str(time.time()
    # minutes_to_add = pandas.Timedelta(minutes=5)
    # time = time + minutes_to_add
    halls_names = config[HALLS_NAMES]
    halls_count = config[HALLS_COUNT]
    slot_size = config[SLOT_SIZE]

    for i_hall in range(0, halls_count):
        hall_name = halls_names[i_hall]
        slots = slots_by_hall_name[hall_name]
        slots_count = len(slots)
        for i_slot in range(0, slots_count):
            start_row = i_slot * (slot_size + 2)  # Plus 2 because we need 2 rows for the header and footer of each slot
            start_col = i_hall * 5  # Multiply by 5, because each slot is wide 4 columns. This makes 1 empty column to the right of each slot

            # Write header
            worksheet.write(start_row, start_col, "#")
            worksheet.write(start_row, start_col + 1, "Ученик")
            worksheet.write(start_row, start_col + 2, "Ментор")
            worksheet.write(start_row, start_col + 3, "Координатор")

            # Write teams
            teams = slots[i_slot].teams
            teams_count = len(teams)
            for i_team in range(teams_count):
                team = teams[i_team]
                worksheet.write(start_row + i_team + 1, start_col, team.number)
                worksheet.write(start_row + i_team + 1, start_col + 1, team.student_name)
                worksheet.write(start_row + i_team + 1, start_col + 2, team.mentor_name)
                worksheet.write(start_row + i_team + 1, start_col + 3, team.coordinator_name)


def create_schedule(season: str):
    config = get_config(season)
    random.seed(config[RANDOM_SEED])

    csv_data = pandas.read_csv(TEAMS_FILE_PATH)
    teams = get_teams(config, csv_data)
    slots_by_hall_name = create_slots(config, teams)
    workbook = xlsxwriter.Workbook(SCHEDULE_FILE_NAME.format(season).lower())
    worksheet = workbook.add_worksheet("Schedule")

    populate_sheet(config, worksheet, slots_by_hall_name)

    worksheet.autofit()
    workbook.close()


if __name__ == "__main__":
    create_schedule(SEASON_SOFIA)
    create_schedule(SEASON_ONLINE)
