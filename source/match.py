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
STUDENT_SYNTHESIZED_INFO = get_column_index("BR")
STUDENT_SURVEY_0 = get_column_index("BJ")
STUDENT_SURVEY_1 = get_column_index("BK")
STUDENT_SURVEY_2 = get_column_index("BL")
STUDENT_SURVEY_3 = get_column_index("BM")
STUDENT_SURVEY_4 = get_column_index("BN")

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
MENTOR_SURVEY_0 = get_column_index("BJ")
MENTOR_SURVEY_1 = get_column_index("BK")
MENTOR_SURVEY_2 = get_column_index("BL")
MENTOR_SURVEY_3 = get_column_index("BM")
MENTOR_SURVEY_4 = get_column_index("BN")

# Comparison weights
INTERESTS_WEIGHT = 5
HOBBIES_WEIGHT = 2.5
PROJECT_TYPE_WEIGHT = 2.5
HOURS_PER_WEEK_WEIGHT = 1
SYNTHESIZED_INFO_WEIGHT = 2
SURVEY_0_WEIGHT = 0.5
SURVEY_1_WEIGHT = 0.5
SURVEY_2_WEIGHT = 0.5
SURVEY_3_WEIGHT = 0.5
SURVEY_4_WEIGHT = 0.5

# Similarities below this percentage will be discarded
SIMILARITY_PERCENT_DISCARD_THRESHOLD = 50

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
        self.synthesized_info = None
        self.survey_0 = None
        self.survey_1 = None
        self.survey_2 = None
        self.survey_3 = None
        self.survey_4 = None
        self._encoded_interests = None
        self._encoded_hobbies = None
        self._encoded_project_type = None
        self._encoded_synthesized_info = None
        self._encoded_survey_0 = None
        self._encoded_survey_1 = None
        self._encoded_survey_2 = None
        self._encoded_survey_3 = None
        self._encoded_survey_4 = None

    def __repr__(self):
        result = f"Name: {self.name}{os.linesep}"
        result += f"Status: {self.status}{os.linesep}"
        result += f"Interests: {self.interests}{os.linesep}"
        result += f"Hobbies: {self.hobbies}{os.linesep}"
        result += f"Project Type: {self.project_type}{os.linesep}"
        result += f"Hours/Week: {self.hours_per_week}{os.linesep}"
        result += f"Synthesized Info: {self.synthesized_info}{os.linesep}"
        result += f"Survey_0: {self.survey_0}{os.linesep}"
        result += f"Survey_1: {self.survey_1}{os.linesep}"
        result += f"Survey_2: {self.survey_2}{os.linesep}"
        result += f"Survey_3: {self.survey_3}{os.linesep}"
        result += f"Survey_4: {self.survey_4}{os.linesep}"
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

    def synthesized_info_enc(self, model: SentenceTransformer):
        if self._encoded_synthesized_info is None:
            self._encoded_synthesized_info = model.encode(self._encoded_synthesized_info)
        return self._encoded_synthesized_info

    def survey_0_enc(self, model: SentenceTransformer):
        if self._encoded_survey_0 is None:
            self._encoded_survey_0 = model.encode(self._encoded_survey_0)
        return self._encoded_survey_0

    def survey_1_enc(self, model: SentenceTransformer):
        if self._encoded_survey_1 is None:
            self._encoded_survey_1 = model.encode(self._encoded_survey_1)
        return self._encoded_survey_1

    def survey_2_enc(self, model: SentenceTransformer):
        if self._encoded_survey_2 is None:
            self._encoded_survey_2 = model.encode(self._encoded_survey_2)
        return self._encoded_survey_2

    def survey_3_enc(self, model: SentenceTransformer):
        if self._encoded_survey_3 is None:
            self._encoded_survey_3 = model.encode(self._encoded_survey_3)
        return self._encoded_survey_3

    def survey_4_enc(self, model: SentenceTransformer):
        if self._encoded_survey_4 is None:
            self._encoded_survey_4 = model.encode(self._encoded_survey_4)
        return self._encoded_survey_4


class PersonDataFilter:
    def __init__(self):
        self.name_index = None
        self.status_index = None
        self.confirmed_index = None
        self.interests_indices = None
        self.hobbies_indices = None
        self.project_type_index = None
        self.hours_per_week_index = None
        self.synthesized_info_indices = None
        self.survey_0_index = None
        self.survey_1_index = None
        self.survey_2_index = None
        self.survey_3_index = None
        self.survey_4_index = None

        self.is_student = False

    def get_max_index(self):
        max_index = max(
            self.name_index,
            self.status_index,
            self.confirmed_index,
            *self.interests_indices,
            *self.hobbies_indices,
            self.project_type_index,
            self.hours_per_week_index,
            *self.synthesized_info_indices,
            self.survey_0_index,
            self.survey_1_index,
            self.survey_2_index,
            self.survey_3_index,
            self.survey_4_index)
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
            synthesized_info = ""
            survey_0 = ""
            survey_1 = ""
            survey_2 = ""
            survey_3 = ""
            survey_4 = ""
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

                if col_i == filter.hours_per_week_index:
                    hours_per_week = parse_hours_per_week(row[col_i])
                    continue

                if col_i in filter.synthesized_info_indices:
                    synthesized_info += f"{row[col_i]}{os.linesep}"
                    continue

                if col_i == filter.survey_0_index:
                    survey_0 = row[col_i]
                    continue

                if col_i == filter.survey_1_index:
                    survey_1 = row[col_i]
                    continue

                if col_i == filter.survey_2_index:
                    survey_2 = row[col_i]
                    continue

                if col_i == filter.survey_3_index:
                    survey_3 = row[col_i]
                    continue

                if col_i == filter.survey_4_index:
                    survey_4 = row[col_i]
                    continue

                if col_i >= filter.get_max_index():
                    person = PersonData()
                    person.name = name
                    person.status = status
                    person.interests = interests
                    person.hobbies = hobbies
                    person.project_type = project_type
                    person.hours_per_week = hours_per_week
                    person.synthesized_info = synthesized_info
                    person.survey_0 = survey_0
                    person.survey_1 = survey_1
                    person.survey_2 = survey_2
                    person.survey_3 = survey_3
                    person.survey_4 = survey_4
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
    students_filter.synthesized_info_indices = [STUDENT_SYNTHESIZED_INFO]
    students_filter.survey_0_index = STUDENT_SURVEY_0
    students_filter.survey_1_index = STUDENT_SURVEY_1
    students_filter.survey_2_index = STUDENT_SURVEY_2
    students_filter.survey_3_index = STUDENT_SURVEY_3
    students_filter.survey_4_index = STUDENT_SURVEY_4
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
    mentors_filter.synthesized_info_indices = [MENTOR_PROFESIONAL_EXPERIENCE, MENTOR_AREAS_OF_INTEREST]
    mentors_filter.survey_0_index = MENTOR_SURVEY_0
    mentors_filter.survey_1_index = MENTOR_SURVEY_1
    mentors_filter.survey_2_index = MENTOR_SURVEY_2
    mentors_filter.survey_3_index = MENTOR_SURVEY_3
    mentors_filter.survey_4_index = MENTOR_SURVEY_4
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
            final_sim = 0
            max_sim = 0

            interests_sim = util.cos_sim(student.interests_enc(model), mentor.interests_enc(model)).item() * INTERESTS_WEIGHT
            final_sim += interests_sim
            max_sim += INTERESTS_WEIGHT

            hobbies_sim = util.cos_sim(student.hobbies_enc(model), mentor.hobbies_enc(model)).item() * HOBBIES_WEIGHT
            final_sim += hobbies_sim
            max_sim += HOBBIES_WEIGHT

            project_type_sim = util.cos_sim(student.project_type_enc(model), mentor.project_type_enc(model)).item() * PROJECT_TYPE_WEIGHT
            final_sim += project_type_sim
            max_sim += PROJECT_TYPE_WEIGHT

            hours_per_week_sim = (1 - abs(student.hours_per_week - mentor.hours_per_week) / (MAX_HOURS_PER_WEEK - MIN_HOURS_PER_WEEK)) * HOURS_PER_WEEK_WEIGHT
            final_sim += hours_per_week_sim
            max_sim += HOURS_PER_WEEK_WEIGHT

            # if student.synthesized_info.strip() != "" and mentor.synthesized_info.strip() != "":
            #     synthesized_info_sim = util.cos_sim(student.synthesized_info_enc(model), mentor.synthesized_info_enc(model)).item() * SYNTHESIZED_INFO_WEIGHT
            #     final_sim += synthesized_info_sim
            #     max_sim += SYNTHESIZED_INFO_WEIGHT

            if student.survey_0.strip() != "" and mentor.survey_0.strip() != "":
                survey_0_sim = util.cos_sim(student.survey_0_enc(model), mentor.survey_0_enc(model)).item() * SURVEY_0_WEIGHT
                final_sim += survey_0_sim
                max_sim += SURVEY_0_WEIGHT

            if student.survey_1.strip() != "" and mentor.survey_1.strip() != "":
                survey_1_sim = util.cos_sim(student.survey_1_enc(model), mentor.survey_1_enc(model)).item() * SURVEY_1_WEIGHT
                final_sim += survey_1_sim
                max_sim += SURVEY_1_WEIGHT

            if student.survey_2.strip() != "" and mentor.survey_2.strip() != "":
                survey_2_sim = util.cos_sim(student.survey_2_enc(model), mentor.survey_2_enc(model)).item() * SURVEY_2_WEIGHT
                final_sim += survey_2_sim
                max_sim += SURVEY_2_WEIGHT

            if student.survey_3.strip() != "" and mentor.survey_3.strip() != "":
                survey_3_sim = util.cos_sim(student.survey_3_enc(model), mentor.survey_3_enc(model)).item() * SURVEY_3_WEIGHT
                final_sim += survey_3_sim
                max_sim += SURVEY_3_WEIGHT

            if student.survey_4.strip() != "" and mentor.survey_4.strip() != "":
                survey_4_sim = util.cos_sim(student.survey_4_enc(model), mentor.survey_4_enc(model)).item() * SURVEY_4_WEIGHT
                final_sim += survey_4_sim
                max_sim += SURVEY_4_WEIGHT

            sim_percent = int((final_sim / max_sim) * 100)
            mentor_suggestions.append((mentor, sim_percent))

        mentor_suggestions.sort(key=lambda x: x[1], reverse=True)
        for mentor_suggestion in mentor_suggestions:
            if mentor_suggestion[1] < SIMILARITY_PERCENT_DISCARD_THRESHOLD:
                break
            match_entry = f"{student.name}(Y) + {mentor_suggestion[0].name}(M) - {mentor_suggestion[1]}"
            match_entry = f"{student.name}"
            match_entry = f"{mentor_suggestion[0].name}"
            match_entry = f"{mentor_suggestion[1]}"
            matches.append(match_entry)
            # print(match_entry)

    with open("matches.txt", "w", encoding="utf-8") as file:
        for entry in matches:
            file.write(f"{entry}\n")


if __name__ == "__main__":
    find_matches()
