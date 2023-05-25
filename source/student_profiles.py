import os
import csv
import docx


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles"
REGISTER_FILE_NAME = "student_register.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}"

STATUS = 0
STUDENT_NAME = 1
AGE = 2
GENDER = 3
SCHOOL_NAME = 4
FORMAT = 5
GRADE = 6
SCHOOL_INTERESTS = 7
NON_SCHOOL_INTERESTS = 8
ENGLISH_LEVEL = 9
SPORT = 10
WHAT_TO_DO_AFTER_SCHOOL = 11
INTERESTS = 12
SKILLS_TO_IMPROVE = 13
FREE_TIME_ACTIVITIES = 14
DIFFICULT_SITUATION = 15
IDEA_IN_ABLE_MENTOR = 16
WANT_TO_CHANGE = 17
HOURS_PER_WEEK = 18
PROJECT_WITH_MENTOR = 19
HEARD_OF_ABLE_MENTOR = 20
STUDENT_NAME_COPY = 21
MENTOR_NAME = 22

column_titles = [None] * 23
column_titles[STATUS] = "Статус"
column_titles[STUDENT_NAME] = "Ученик"
column_titles[AGE] = "Възраст"
column_titles[GENDER] = "Пол"
column_titles[SCHOOL_NAME] = "Училище"
column_titles[FORMAT] = "Населено място"
column_titles[GRADE] = "Завършен клас"
column_titles[SCHOOL_INTERESTS] = "Интереси, свързани с училище"
column_titles[NON_SCHOOL_INTERESTS] = "Интереси извън училище"
column_titles[ENGLISH_LEVEL] = "Ниво на английски език"
column_titles[SPORT] = "Спорт"
column_titles[WHAT_TO_DO_AFTER_SCHOOL] = "Какво ще правя след гимназията:"
column_titles[INTERESTS] = "В кои сфери имаш интерес да се развиваш и в кои по-слаб?"
column_titles[SKILLS_TO_IMPROVE] = "Кои свои качества искаш да промениш/ подобриш?"
column_titles[FREE_TIME_ACTIVITIES] = "Как се забавляваш в свободното си време?"
column_titles[DIFFICULT_SITUATION] = "Разкажи ни за трудна ситуация и как си се справил/а?"
column_titles[IDEA_IN_ABLE_MENTOR] = "Каква идея искаш да осъществиш в рамките на ABLE Mentor?"
column_titles[WANT_TO_CHANGE] = "Желая да променя..."
column_titles[HOURS_PER_WEEK] = "По колко часа седмично би отделял/а на проекта?"
column_titles[PROJECT_WITH_MENTOR] = "По какъв проект би работил/а със своя ментор?"
column_titles[HEARD_OF_ABLE_MENTOR] = "Научил/а за ABLE Mentor от?"
column_titles[STUDENT_NAME_COPY] = "Ученик"
column_titles[MENTOR_NAME] = "Ментор"


def try_create_doc(row_data, file_path):
    if row_data[STATUS] != "Активен":
        return False

    doc = docx.Document()
    doc.add_heading("ПРОФИЛ НА ТВОЯ УЧЕНИК", 0)

    table = doc.add_table(rows=0, cols=2)

    for id, entry in enumerate(row_data):
        if (id == STATUS or
            id == STUDENT_NAME or
            id == GENDER or
            id == FORMAT or
            id == STUDENT_NAME_COPY or
                id == MENTOR_NAME):
            continue

        table_row = table.add_row().cells
        table_row[0].text = column_titles[id]
        table_row[1].text = row_data[id]

    doc.save(file_path)
    return True


def create_docs():
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.mkdir(OUTPUT_DIRECTORY)

    with open(REGISTER_FILE_PATH, encoding="utf-8", mode="r") as fstream:
        reader = csv.reader(fstream, delimiter=',', quotechar='"')

        for idx, row in enumerate(reader):
            if idx == 0:
                continue  # skip first row

            mentor_name = row[MENTOR_NAME].replace("/", "").strip()
            file_path = f"{OUTPUT_DIRECTORY}/{mentor_name}.docx".replace("\\", "/")
            try_create_doc(row, file_path)


if __name__ == "__main__":
    create_docs()
