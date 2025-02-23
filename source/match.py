import csv
import math
import os
from sentence_transformers import SentenceTransformer, util


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
STUDENTS_CSV_FILE_NAME = "match_data_students.csv"
STUDENTS_CSV_FILE_PATH = f"{CURRENT_DIRECTORY}/{STUDENTS_CSV_FILE_NAME}"
MENTORS_CSV_FILE_NAME = "match_data_mentors.csv"
MENTORS_CSV_FILE_PATH = f"{CURRENT_DIRECTORY}/{MENTORS_CSV_FILE_NAME}"
AI_MODEL_FILE_PATH = f"{CURRENT_DIRECTORY}/ai_models/paraphrase-multilingual-MiniLM-L12-v2"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


# Column indices in the students' csv file
STUDENT_NAME = get_column_index("A")
STUDENT_STATUS = get_column_index("S")
STUDENT_CONFIRMED = get_column_index("T")
STUDENT_NON_SCHOOL_INTERESTS = get_column_index("AK")  # interests
STUDENT_AREAS_OF_INTEREST = get_column_index("AQ")  # interests
STUDENT_MENTOR_PROFESIONAL_EXPERIENCE = get_column_index("AR")  # interests
STUDENT_SPORT = get_column_index("AO")  # hobbies
STUDENT_HOBBIES = get_column_index("AT")  # hobbies
STUDENT_PROJECT_TYPE = get_column_index("AZ")
STUDENT_HOURS_PER_WEEK = get_column_index("AY")

# Column indices in the mentors' csv file
MENTOR_NAME = get_column_index("I")
MENTOR_STATUS = get_column_index("A")
MENTOR_CONFIRMED = get_column_index("B")
MENTOR_EDUCATION = get_column_index("Q")  # interests
MENTOR_PROFESIONAL_EXPERIENCE = get_column_index("S")  # interests
MENTOR_AREAS_OF_INTEREST = get_column_index("U")  # interests
MENTOR_HOBBIES = get_column_index("V")  # hobbies
MENTOR_PROJECT_TYPE = get_column_index("Z")
MENTOR_HOURS_PER_WEEK = get_column_index("Y")

# Comparison weights
INTERESTS_WEIGHT = 4
HOBBIES_WEIGHT = 2
PROJECT_TYPE_WEIGHT = 2
HOURS_PER_WEEK_WEIGHT = 1
SIMILARITY_PERCENT_DISCARD_THRESHOLD = 50  # similarities below this percentage will be discarded

# Hours per week
MIN_HOURS_PER_WEEK = 1
MAX_HOURS_PER_WEEK = 6


class PersonData:
    def __init__(self):
        self.name = None
        self.status = None
        self.interests = None
        self.hobbies = None
        self.project_type = None
        self.hours_per_week = None
        self._encoded_interests = None
        self._encoded_hobbies = None
        self._encoded_project_type = None

    def __repr__(self):
        result = f"Name: {self.name}{os.linesep}"
        result += f"Status: {self.status}{os.linesep}"
        result += f"Interests: {self.interests}{os.linesep}"
        result += f"Hobbies: {self.hobbies}{os.linesep}"
        result += f"Project Type: {self.project_type}{os.linesep}"
        result += f"Hours/Week: {self.hours_per_week}{os.linesep}"
        return result

    def interests_enc(self, model: SentenceTransformer):
        if self._encoded_interests is None:
            self._encoded_interests = model.encode(self.interests)
        return self._encoded_interests

    def hobbies_enc(self, model: SentenceTransformer):
        if self._encoded_hobbies is None:
            self._encoded_hobbies = model.encode(self.hobbies)
        return self._encoded_hobbies

    def project_type_enc(self, model: SentenceTransformer):
        if self._encoded_project_type is None:
            self._encoded_project_type = model.encode(self.project_type)
        return self._encoded_project_type


class PersonDataFilter:
    def __init__(self):
        self.name_index = None
        self.status_index = None
        self.confirmed_index = None
        self.interests_indices = None
        self.hobbies_indices = None
        self.project_type_index = None
        self.hours_per_week_index = None
        self.is_student = False

    def get_max_index(self):
        max_index = max(self.name_index, self.status_index, self.confirmed_index, *self.interests_indices,
                        *self.hobbies_indices, self.project_type_index, self.hours_per_week_index)
        return max_index


def parse_hours_per_week(hours_per_week: str):
    try:
        return int(hours_per_week)
    except ValueError:
        return MAX_HOURS_PER_WEEK


def extract_people_data(csv_file_path: str, filter: PersonDataFilter) -> list[PersonData]:
    with open(csv_file_path, encoding="utf-8", mode="r") as fstream:
        people = list()
        reader = csv.reader(fstream, delimiter=',', quotechar='"')
        for row_i, row in enumerate(reader):
            if row_i == 0:
                continue

            name = row[filter.name_index]
            if not name:
                continue

            status = row[filter.status_index].lower()
            if filter.is_student:
                if status == "matched" or status.startswith("отпаднал"):
                    continue
            else:
                if status == "matched" or status == "no matching!":
                    continue

            confirmed = row[filter.confirmed_index].lower()
            if filter.is_student:
                if confirmed != "да":
                    continue
            else:
                if confirmed == "denied":
                    continue

            interests = ""
            hobbies = ""
            project_type = ""
            hours_per_week = 0
            for col_i in range(0, len(row)):
                if col_i in filter.interests_indices:
                    interests += f"{row[col_i]}{os.linesep}"

                if col_i in filter.hobbies_indices:
                    hobbies += f"{row[col_i]}{os.linesep}"

                if col_i == filter.project_type_index:
                    project_type = row[col_i]

                if col_i == filter.hours_per_week_index:
                    hours_per_week = parse_hours_per_week(row[col_i])

                if col_i >= filter.get_max_index():
                    person = PersonData()
                    person.name = name
                    person.status = status
                    person.interests = interests
                    person.hobbies = hobbies
                    person.project_type = project_type
                    person.hours_per_week = hours_per_week
                    people.append(person)
                    break

        people.sort(key=lambda p: p.name)
        return people


def find_matches():
    students_filter = PersonDataFilter()
    students_filter.name_index = STUDENT_NAME
    students_filter.status_index = STUDENT_STATUS
    students_filter.confirmed_index = STUDENT_CONFIRMED
    students_filter.interests_indices = [STUDENT_NON_SCHOOL_INTERESTS, STUDENT_AREAS_OF_INTEREST, STUDENT_MENTOR_PROFESIONAL_EXPERIENCE]
    students_filter.hobbies_indices = [STUDENT_SPORT, STUDENT_HOBBIES]
    students_filter.project_type_index = STUDENT_PROJECT_TYPE
    students_filter.hours_per_week_index = STUDENT_HOURS_PER_WEEK
    students_filter.is_student = True
    students = extract_people_data(STUDENTS_CSV_FILE_PATH, students_filter)

    mentors_filter = PersonDataFilter()
    mentors_filter.name_index = MENTOR_NAME
    mentors_filter.status_index = MENTOR_STATUS
    mentors_filter.confirmed_index = MENTOR_CONFIRMED
    mentors_filter.interests_indices = [MENTOR_EDUCATION, MENTOR_PROFESIONAL_EXPERIENCE, MENTOR_AREAS_OF_INTEREST]
    mentors_filter.hobbies_indices = [MENTOR_HOBBIES]
    mentors_filter.project_type_index = MENTOR_PROJECT_TYPE
    mentors_filter.hours_per_week_index = MENTOR_HOURS_PER_WEEK
    mentors_filter.is_student = False
    mentors = extract_people_data(MENTORS_CSV_FILE_PATH, mentors_filter)

    model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    # model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-mpnet-base-v2")
    # model = SentenceTransformer(AI_MODEL_FILE_PATH)

    print("find matches:")
    matches = list()
    for student in students:
        mentor_suggestions = list()
        for mentor in mentors:
            # find similarity
            interests_sim = util.cos_sim(student.interests_enc(model), mentor.interests_enc(model)).item() * INTERESTS_WEIGHT
            hobbies_sim = util.cos_sim(student.hobbies_enc(model), mentor.hobbies_enc(model)).item() * HOBBIES_WEIGHT
            project_type_sim = util.cos_sim(student.project_type_enc(model), mentor.project_type_enc(model)).item() * PROJECT_TYPE_WEIGHT
            hours_per_week_sim = (1 - abs(student.hours_per_week - mentor.hours_per_week) / (MAX_HOURS_PER_WEEK - MIN_HOURS_PER_WEEK)) * HOURS_PER_WEEK_WEIGHT
            final_sim = interests_sim + hobbies_sim + project_type_sim + hours_per_week_sim
            max_sim = INTERESTS_WEIGHT + HOBBIES_WEIGHT + PROJECT_TYPE_WEIGHT + HOURS_PER_WEEK_WEIGHT
            sim_percent = int((final_sim / max_sim) * 100)
            mentor_suggestions.append((mentor, sim_percent))

        mentor_suggestions.sort(key=lambda x: x[1], reverse=True)
        for mentor_suggestion in mentor_suggestions:
            if mentor_suggestion[1] < SIMILARITY_PERCENT_DISCARD_THRESHOLD:
                break
            match_entry = f"{student.name}(Y) + {mentor_suggestion[0].name}(M) - {mentor_suggestion[1]}"
            matches.append(match_entry)
            print(match_entry)

    with open("matches.txt", "w", encoding="utf-8") as file:
        for entry in matches:
            file.write(f"{entry}\n")


if __name__ == "__main__":
    find_matches()
