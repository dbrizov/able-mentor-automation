import csv
import math
import os
from sentence_transformers import SentenceTransformer, util


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
STUDENTS_DATA_FILE_NAME = "match_data_students.csv"
STUDENTS_DATA_FILE_PATH = f"{CURRENT_DIRECTORY}/{STUDENTS_DATA_FILE_NAME}"
MENTORS_DATA_FILE_NAME = "match_data_mentors.csv"
MENTORS_DATA_FILE_PATH = f"{CURRENT_DIRECTORY}/{MENTORS_DATA_FILE_NAME}"
AI_MODEL_FILE_PATH = f"{CURRENT_DIRECTORY}/ai_models/paraphrase-multilingual-MiniLM-L12-v2"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


STUDENT_NAME_INDEX = get_column_index("D")
MENTOR_NAME_INDEX = get_column_index("A")

STUDENT_COLUMN_INDICES = [
    get_column_index("AR"),  # Интереси, свързани с училище
    get_column_index("AS"),  # Интереси извън училище
    get_column_index("AU"),  # Каква е нуждата
    get_column_index("AW"),  # Спорт
    get_column_index("AX"),  # Какво ще правя след гимназията
    get_column_index("AY"),  # Сфери, които са ти интересни и в които искаш да се развиваш
    get_column_index("AZ"),  # Ментор в каква професионална сфера би бил/а най-полезен/а за теб
    get_column_index("BA"),  # Кои свои качества искаш да промениш/ подобриш
    get_column_index("BB"),  # Как се забавляваш в свободното си време
    get_column_index("BD"),  # Защо кандидатстваш
    get_column_index("BE"),  # Каква идея искаш да осъществиш в рамките на ABLE Mentor
    get_column_index("BF"),  # Желая да променя...
    get_column_index("BH"),  # По какъв проект би работил/а със своя ментор
    # get_column_index("BK"),  # Впечатления от инфо среща /характер, проект, ментор/
]


MENTOR_COLUMN_INDICES = [
    get_column_index("O"),  # Образование – специалност (посочете всички специалности, ако имате повече от една)?
    get_column_index("P"),  # Име на учебното заведение, от където е придобита последната образователна степен:
    get_column_index("Q"),  # Къде работите в момента (организация/компания)?
    get_column_index("R"),  # Сфери, в които работите и/или сте работили досега (професионален опит):
    get_column_index("S"),  # Сфери, в които имате опит и интереси?
    get_column_index("T"),  # Разкажете ни за Вашите интереси/хобита/компетенции, различни от професионалните Ви такива? Какъв е опитът Ви в тези сфери?
    get_column_index("V"),  # Желая да променя/подобря...
    get_column_index("X"),  # По какъв проект бихте работили със своя ученик?
]


class Person:
    def __init__(self, name: str, data: str):
        self.name = name
        self.data = data
        self._encoded_data = None

    def __repr__(self):
        return f"({self.name}, {self.data})"

    def encoded_data(self, model):
        if self._encoded_data is None:
            self._encoded_data = model.encode(self.data)
        return self._encoded_data


def extract_people_data(csv_file_path: str, csv_column_indices: list, person_name_column_index: int):
    with open(csv_file_path, encoding="utf-8", mode="r") as fstream:
        people = list()
        reader = csv.reader(fstream, delimiter=',', quotechar='"')
        for row_i, row in enumerate(reader):
            if row_i == 0:
                continue

            person_name = row[person_name_column_index]
            if not person_name:
                continue

            person_data = ""
            for col_i in range(0, len(row)):
                if col_i in csv_column_indices:
                    person_data += f"{row[col_i]}{os.linesep}"

                if col_i >= csv_column_indices[-1]:
                    person = Person(person_name, person_data)
                    people.append(person)
                    break

        return people


def find_matches():
    print("extract students")
    students = extract_people_data(STUDENTS_DATA_FILE_PATH, STUDENT_COLUMN_INDICES, STUDENT_NAME_INDEX)
    print("extract mentors")
    mentors = extract_people_data(MENTORS_DATA_FILE_PATH, MENTOR_COLUMN_INDICES, MENTOR_NAME_INDEX)
    print("create model")
    model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    # model = SentenceTransformer(AI_MODEL_FILE_PATH)

    print("find matches:")
    for student in students:
        max_similarity = -1
        matched_mentor = None
        for mentor in mentors:
            similarity = util.cos_sim(student.encoded_data(model), mentor.encoded_data(model)).item()
            if similarity > max_similarity:
                max_similarity = similarity
                matched_mentor = mentor

        print(f"{student.name}(Y) + {matched_mentor.name}(M) - {max_similarity:.2f}")


if __name__ == "__main__":
    find_matches()
