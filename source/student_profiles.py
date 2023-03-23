import os
import csv
import docx


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__))
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/StudentProfiles"
REGISTER_FILE_NAME = "student_register.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}".replace("\\", "/")

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

texts = [None] * 23
texts[STATUS] = "Статус"
texts[STUDENT_NAME] = "Ученик"
texts[AGE] = "Възраст"
texts[GENDER] = "Пол"
texts[SCHOOL_NAME] = "Училище"
texts[FORMAT] = "Населено място"
texts[GRADE] = "Завършен клас"
texts[SCHOOL_INTERESTS] = "Интереси, свързани с училище"
texts[NON_SCHOOL_INTERESTS] = "Интереси извън училище"
texts[ENGLISH_LEVEL] = "Ниво на английски език"
texts[SPORT] = "Спорт"
texts[WHAT_TO_DO_AFTER_SCHOOL] = "Какво ще правя след гимназията:"
texts[INTERESTS] = "В кои сфери имаш интерес да се развиваш и в кои по-слаб?"
texts[SKILLS_TO_IMPROVE] = "Кои свои качества искаш да промениш/ подобриш?"
texts[FREE_TIME_ACTIVITIES] = "Как се забавляваш в свободното си време?"
texts[DIFFICULT_SITUATION] = "Разкажи ни за трудна ситуация и как си се справил/а?"
texts[IDEA_IN_ABLE_MENTOR] = "Каква идея искаш да осъществиш в рамките на ABLE Mentor?"
texts[WANT_TO_CHANGE] = "Желая да променя..."
texts[HOURS_PER_WEEK] = "По колко часа седмично би отделял/а на проекта?"
texts[PROJECT_WITH_MENTOR] = "По какъв проект би работил/а със своя ментор?"
texts[HEARD_OF_ABLE_MENTOR] = "Научил/а за ABLE Mentor от?"
texts[STUDENT_NAME_COPY] = "Ученик"
texts[MENTOR_NAME] = "Ментор"


def try_create_doc(register_row, file_path):
    if (register_row[STATUS] != "Активен"):
        return False

    doc = docx.Document()
    doc.add_heading("ПРОФИЛ НА ТВОЯ УЧЕНИК", 0)

    table = doc.add_table(rows=0, cols=2)

    for id, entry in enumerate(register_row):
        if (id == STATUS or
            id == STUDENT_NAME or
            id == GENDER or
            id == FORMAT or
            id == STUDENT_NAME_COPY or
                id == MENTOR_NAME):
            continue

        row = table.add_row().cells
        row[0].text = texts[id]
        row[1].text = register_row[id]

    doc.save(file_path)
    return True


def create_docs():
    if (not os.path.exists(OUTPUT_DIRECTORY)):
        os.mkdir(OUTPUT_DIRECTORY)

    with open(REGISTER_FILE_PATH, encoding="utf-8", mode="r") as fstream:
        reader = csv.reader(fstream, delimiter=',', quotechar='"')

        for idx, row in enumerate(reader):
            if (idx == 0):
                continue  # skip first row

            mentor_name = row[MENTOR_NAME].replace("/", "")
            file_path = f"{OUTPUT_DIRECTORY}/{mentor_name}.docx".replace("\\", "/")
            try_create_doc(row, file_path)


create_docs()
