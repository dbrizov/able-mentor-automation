import math
import os
import pandas
import json
import random
import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
CONFIG_FILE_NAME = "schedule.json"
CONFIG_FILE_PATH = f"{CURRENT_DIRECTORY}/{CONFIG_FILE_NAME}"
TEAMS_FILE_NAME = "teams.csv"
TEAMS_FILE_PATH = f"{CURRENT_DIRECTORY}/{TEAMS_FILE_NAME}"
SCHEDULE_FILE_NAME = "schedule_{0}.xlsx"

# Config keys
CSV = "csv"
ACTIVE_COLUMN = "active_column"
STUDENT_COLUMN = "student_column"
MENTOR_COLUMN = "mentor_column"
COORDINATOR_COLUMN = "coordinator_column"
SEASON_TYPE_COLUMN = "season_type_column"

FORMAT = "format"
ROW_HEIGHT = "row_height"
COLUMN_WIDTH = "column_width"
ROWS_BETWEEN_SLOTS = "rows_between_slots"
COLUMNS_BETWEEN_SLOTS = "columns_between_slots"

SOFIA = "sofia"
ONLINE = "online"
SEASON_TYPE = "season_type"
SLOT_SIZE = "slot_size"
HALLS = "halls"
NAME = "name"
COORDINATORS = "coordinators"
RANDOM_SEED = "random_seed"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


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
    with open(CONFIG_FILE_PATH, mode="r", encoding="utf-8") as file_stream:
        js = json.load(file_stream)
        config = {
            # csv
            ACTIVE_COLUMN: js[CSV][ACTIVE_COLUMN],
            STUDENT_COLUMN: js[CSV][STUDENT_COLUMN],
            MENTOR_COLUMN: js[CSV][MENTOR_COLUMN],
            COORDINATOR_COLUMN: js[CSV][COORDINATOR_COLUMN],
            SEASON_TYPE_COLUMN: js[CSV][SEASON_TYPE_COLUMN],
            # format
            ROW_HEIGHT: js[FORMAT][ROW_HEIGHT],
            COLUMN_WIDTH: js[FORMAT][COLUMN_WIDTH],
            ROWS_BETWEEN_SLOTS: js[FORMAT][ROWS_BETWEEN_SLOTS],
            COLUMNS_BETWEEN_SLOTS: js[FORMAT][COLUMNS_BETWEEN_SLOTS],
            # season
            SEASON_TYPE: js[season][SEASON_TYPE],
            SLOT_SIZE: js[season][SLOT_SIZE],
            HALLS: js[season][HALLS],
            RANDOM_SEED: js[season][RANDOM_SEED],
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
    active_column_index = get_column_index(config[ACTIVE_COLUMN])
    student_column_index = get_column_index(config[STUDENT_COLUMN])
    mentor_column_index = get_column_index(config[MENTOR_COLUMN])
    coordinator_column_index = get_column_index(config[COORDINATOR_COLUMN])
    season_type_column_index = get_column_index(config[SEASON_TYPE_COLUMN])

    active = get_column_data(csv_data, active_column_index)
    students = get_column_data(csv_data, student_column_index)
    mentors = get_column_data(csv_data, mentor_column_index)
    coordinators = get_column_data(csv_data, coordinator_column_index)
    season_types = get_column_data(csv_data, season_type_column_index)

    teams = list()
    team_number = 0
    for i in range(len(active)):
        if active[i] == "Активен" and season_types[i] == config[SEASON_TYPE]:
            teams.append(Team(team_number, students[i], mentors[i], coordinators[i]))

    return teams


def create_slots(config, teams: list):
    # Group teams by coordinator
    teams_by_coordinator = dict()
    for team in teams:
        coordinator_name = team.coordinator_name.lower()
        if coordinator_name not in teams_by_coordinator:
            teams_by_coordinator[coordinator_name] = list()
        teams_by_coordinator[coordinator_name].append(team)

    # Move the halls' specified coordinators to the end of the coordinators' list
    halls = config[HALLS]
    coordinators_names = list(teams_by_coordinator.keys())
    random.shuffle(coordinators_names)
    for hall in halls:
        coordinators = [x.lower() for x in hall[COORDINATORS]]
        for coordinator in coordinators:
            if coordinator in coordinators_names:
                coordinators_names.remove(coordinator)
                coordinators_names.append(coordinator)

    # Group teams by hall
    last_coordinator_index = -1
    coordinators_count = len(coordinators_names)
    teams_by_hall_name = dict()
    for hall in halls:
        # Move this hall's specified coordinators to the start of the coordinators' list
        # This will ensure that these coordinators will belong to this hall
        coordinators = [x.lower() for x in hall[COORDINATORS]]
        for coordinator in coordinators:
            if coordinator in coordinators_names:
                coordinators_names.remove(coordinator)
                coordinators_names.insert(last_coordinator_index + 1, coordinator)

        teams_in_hall = list()
        for i_coordinator in range(last_coordinator_index + 1, coordinators_count):
            last_coordinator_index = i_coordinator
            coordinator_name = coordinators_names[i_coordinator]
            teams_in_hall += teams_by_coordinator[coordinator_name]
            teams_in_hall_count = len(teams_in_hall)
            teams_count_per_hall = math.ceil(len(teams) / len(halls))
            if teams_in_hall_count >= teams_count_per_hall or i_coordinator == coordinators_count - 1:
                random.shuffle(teams_in_hall)
                hall_name = hall[NAME]
                teams_by_hall_name[hall_name] = teams_in_hall
                break

    # Create slots by hall
    slot_size = config[SLOT_SIZE]
    slots_by_hall_name = dict()
    for hall in halls:
        hall_name = hall[NAME]
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


def populate_sheet(config, workbook: Workbook, worksheet: Worksheet, slots_by_hall_name: dict):
    # time = pandas.to_datetime("11:00:00")  # str(time.time())
    # minutes_to_add = pandas.Timedelta(minutes=5)
    # time = time + minutes_to_add
    normal_format = workbook.add_format({
        "border": 1,
        "border_color": "black",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True
    })

    bold_format = workbook.add_format({
        "border": 1,
        "border_color": "black",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "bold": True
    })

    header_format = workbook.add_format({
        "border": 1,
        "border_color": "black",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "bold": True,
        "bg_color": "#DCE5F2"
    })

    halls = config[HALLS]
    halls_count = len(halls)
    slot_size = config[SLOT_SIZE]
    slot_header_rows = 2
    slot_columns = 4
    rows_between_slots = config[ROWS_BETWEEN_SLOTS]
    columns_between_slots = config[COLUMNS_BETWEEN_SLOTS]
    row_height = config[ROW_HEIGHT]
    column_width = config[COLUMN_WIDTH]

    for i_hall in range(0, halls_count):
        hall_name = halls[i_hall][NAME]
        slots = slots_by_hall_name[hall_name]
        slots_count = len(slots)
        for i_slot in range(0, slots_count):
            start_row = i_slot * (slot_size + slot_header_rows + rows_between_slots)
            start_col = i_hall * (slot_columns + columns_between_slots)

            # Resize columns
            worksheet.set_column(start_col, start_col + 3, column_width)

            # Write header
            worksheet.merge_range(start_row, start_col, start_row, start_col + 3, f"Зала {hall_name} (ЧАСТ {i_slot + 1})", header_format)
            row = start_row + 1
            worksheet.write(row, start_col, "#", bold_format)
            worksheet.write(row, start_col + 1, "Ученик", bold_format)
            worksheet.write(row, start_col + 2, "Ментор", bold_format)
            worksheet.write(row, start_col + 3, "Координатор", bold_format)

            # Write teams
            teams = slots[i_slot].teams
            teams_count = len(teams)
            for i_team in range(teams_count):
                team = teams[i_team]
                row = start_row + i_team + 2
                worksheet.set_row(row, row_height)
                worksheet.write(row, start_col, team.number, normal_format)
                worksheet.write(row, start_col + 1, team.student_name, normal_format)
                worksheet.write(row, start_col + 2, team.mentor_name, normal_format)
                worksheet.write(row, start_col + 3, team.coordinator_name, normal_format)


def create_schedule(season: str):
    config = get_config(season)
    random.seed(config[RANDOM_SEED])

    csv_data = pandas.read_csv(TEAMS_FILE_PATH)
    teams = get_teams(config, csv_data)
    slots_by_hall_name = create_slots(config, teams)
    workbook = xlsxwriter.Workbook(SCHEDULE_FILE_NAME.format(season).lower())
    worksheet = workbook.add_worksheet("Schedule")
    populate_sheet(config, workbook, worksheet, slots_by_hall_name)
    workbook.close()


if __name__ == "__main__":
    create_schedule(SOFIA)
    create_schedule(ONLINE)
