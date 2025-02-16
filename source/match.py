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
STUDENT_STATUS = get_column_index("Q")
STUDENT_CONFIRMED = get_column_index("R")
STUDENT_NON_SCHOOL_INTERESTS = get_column_index("AJ")  # interests
STUDENT_AREAS_OF_INTEREST = get_column_index("AP")  # interests
STUDENT_MENTOR_PROFESIONAL_EXPERIENCE = get_column_index("AQ")  # interests
STUDENT_SPORT = get_column_index("AN")  # hobbies
STUDENT_HOBBIES = get_column_index("AS")  # hobbies
STUDENT_PROJECT_TYPE = get_column_index("AY")

# Column indices in the mentors' csv file
MENTOR_NAME = get_column_index("H")
MENTOR_STATUS = get_column_index("A")
MENTOR_EDUCATION = get_column_index("P")  # interests
MENTOR_PROFESIONAL_EXPERIENCE = get_column_index("R")  # interests
MENTOR_AREAS_OF_INTEREST = get_column_index("T")  # interests
MENTOR_HOBBIES = get_column_index("U")  # hobbies
MENTOR_PROJECT_TYPE = get_column_index("Y")

# Александър Алексиев (У) + Ралица Костадинова (М)

# Comparison weights
INTERESTS_WEIGHT = 4
HOBBIES_WEIGHT = 2
PROJECT_TYPE_WEIGHT = 1


class PersonData:
    def __init__(self):
        self.name = None
        self.status = None
        self.interests = None
        self.hobbies = None
        self.project_type = None
        self._encoded_interests = None
        self._encoded_hobbies = None
        self._encoded_project_type = None

    def __repr__(self):
        result = f"Name: {self.name}{os.linesep}"
        result += f"Status: {self.status}{os.linesep}"
        result += f"Interests: {self.interests}{os.linesep}"
        result += f"Hobbies: {self.hobbies}{os.linesep}"
        result += f"Project Type: {self.project_type}{os.linesep}"
        return result

    def encoded_interests(self, model: SentenceTransformer):
        if self._encoded_interests is None:
            self._encoded_interests = model.encode(self.interests)
        return self._encoded_interests

    def encoded_hobbies(self, model: SentenceTransformer):
        if self._encoded_hobbies is None:
            self._encoded_hobbies = model.encode(self.hobbies)
        return self._encoded_hobbies

    def encoded_project_type(self, model: SentenceTransformer):
        if self._encoded_project_type is None:
            self._encoded_project_type = model.encode(self.project_type)
        return self._encoded_project_type


class PersonDataFilter:
    def __init__(self):
        self.name_index = None
        self.status_index = None
        self.interests_indices = None
        self.hobbies_indices = None
        self.project_type_index = None

    def get_max_index(self):
        max_index = max(*self.interests_indices, *self.hobbies_indices, self.project_type_index)
        return max_index


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

            status = row[filter.status_index]
            if status == "matched":
                continue

            interests = ""
            hobbies = ""
            project_type = ""
            for col_i in range(0, len(row)):
                if col_i in filter.interests_indices:
                    interests += f"{row[col_i]}{os.linesep}"
                    continue

                if col_i in filter.hobbies_indices:
                    hobbies += f"{row[col_i]}{os.linesep}"
                    continue

                if col_i == filter.project_type_index:
                    project_type = row[col_i]
                    continue

                if col_i >= filter.get_max_index():
                    person = PersonData()
                    person.name = name
                    person.status = status
                    person.interests = interests
                    person.hobbies = hobbies
                    person.project_type = project_type
                    people.append(person)
                    break

        people.sort(key=lambda p: p.name)
        return people


def find_matches():
    students_filter = PersonDataFilter()
    students_filter.name_index = STUDENT_NAME
    students_filter.status_index = STUDENT_STATUS
    students_filter.interests_indices = [STUDENT_NON_SCHOOL_INTERESTS, STUDENT_AREAS_OF_INTEREST, STUDENT_MENTOR_PROFESIONAL_EXPERIENCE]
    students_filter.hobbies_indices = [STUDENT_SPORT, STUDENT_HOBBIES]
    students_filter.project_type_index = STUDENT_PROJECT_TYPE
    students = extract_people_data(STUDENTS_CSV_FILE_PATH, students_filter)

    mentors_filter = PersonDataFilter()
    mentors_filter.name_index = MENTOR_NAME
    mentors_filter.status_index = MENTOR_STATUS
    mentors_filter.interests_indices = [MENTOR_EDUCATION, MENTOR_PROFESIONAL_EXPERIENCE, MENTOR_AREAS_OF_INTEREST]
    mentors_filter.hobbies_indices = [MENTOR_HOBBIES]
    mentors_filter.project_type_index = MENTOR_PROJECT_TYPE
    mentors = extract_people_data(MENTORS_CSV_FILE_PATH, mentors_filter)

    model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    # model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-mpnet-base-v2")
    # model = SentenceTransformer(AI_MODEL_FILE_PATH)

    print("find matches:")
    for student in students:
        mentor_suggestions = list()
        for mentor in mentors:
            # find similarity
            interests_sim = util.cos_sim(student.encoded_interests(model), mentor.encoded_interests(model)).item() * INTERESTS_WEIGHT
            hobbies_sim = util.cos_sim(student.encoded_hobbies(model), mentor.encoded_hobbies(model)).item() * HOBBIES_WEIGHT
            project_type_sim = util.cos_sim(student.encoded_project_type(model), mentor.encoded_project_type(model)).item() * PROJECT_TYPE_WEIGHT
            final_sim = interests_sim + hobbies_sim + project_type_sim
            max_sim = INTERESTS_WEIGHT + HOBBIES_WEIGHT + PROJECT_TYPE_WEIGHT
            sim_percent = final_sim / max_sim
            mentor_suggestions.append((mentor, sim_percent))

        mentor_suggestions.sort(key=lambda x: x[1], reverse=True)
        for i in range(0, 5):
            print(f"{student.name}(Y) + {mentor_suggestions[i][0].name}(M) - {mentor_suggestions[i][1]:.2f}")


if __name__ == "__main__":
    find_matches()
