import os
import csv
import docx
import math


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles"
REGISTER_FILE_NAME = "student_register.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}"


def get_column_index(column):
    # 26 number system where [A...Z] is mapped to [1...26]
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * math.pow(26, len(column) - idx - 1)

    return int(decimal_value - 1)  # the index is the decimal value minus 1


STATUS = get_column_index("C")
CONFIRMED = get_column_index("J")
STUDENT_NAME = get_column_index("Z")
AGE = get_column_index("AE")
CITY = get_column_index("AF")
SCHOOL_NAME = get_column_index("AG")
GRADE = get_column_index("AH")
SCHOOL_INTERESTS = get_column_index("AI")
NON_SCHOOL_INTERESTS = get_column_index("AJ")
SIMILAR_PROGRAMS = get_column_index("AK")
NEED = get_column_index("AL")
ENGLISH_LEVEL = get_column_index("AM")
SPORT = get_column_index("AN")
WHAT_TO_DO_AFTER_SCHOOL = get_column_index("AO")
INTERESTS = get_column_index("AP")
MENTOR_STUDIED = get_column_index("AQ")
MENTOR_EXPERIENCE = get_column_index("AR")
SKILLS_TO_IMPROVE = get_column_index("AS")
FREE_TIME_ACTIVITIES = get_column_index("AT")
DIFFICULT_SITUATION = get_column_index("AU")
WHY_APPLY = get_column_index("AV")
IDEA_IN_ABLE_MENTOR = get_column_index("AW")
WANT_TO_CHANGE = get_column_index("AX")
HOURS_PER_WEEK = get_column_index("AY")
PROJECT_WITH_MENTOR = get_column_index("AZ")
HEARD_OF_ABLE_MENTOR = get_column_index("BA")
STUDENT_NAME_COPY = get_column_index("BR")
MENTOR_NAME = get_column_index("BT")
COORDINATOR_NAME = get_column_index("BX")

column_titles = dict()
column_titles[STATUS] = "Статус"
column_titles[CONFIRMED] = "Потвърдил участие"
column_titles[STUDENT_NAME] = "Ученик"
column_titles[AGE] = "Възраст"
column_titles[CITY] = "Населено място"
column_titles[SCHOOL_NAME] = "Училище"
column_titles[GRADE] = "Завършен клас"
column_titles[SCHOOL_INTERESTS] = "Интереси, свързани с училище"
column_titles[NON_SCHOOL_INTERESTS] = "Интереси извън училище"
column_titles[SIMILAR_PROGRAMS] = "Участвал ли си в други сходни програми, завършил ли си ги и какво си взе от тях?"
column_titles[NEED] = "ABLE Mentor е напълно безплатна програма с ограничен капацитет. Защо имаш нужда да участваш в нея?"
column_titles[ENGLISH_LEVEL] = "Ниво на английски език"
column_titles[SPORT] = "Спорт"
column_titles[WHAT_TO_DO_AFTER_SCHOOL] = "Какво ще правя след гимназията:"
column_titles[INTERESTS] = "В кои сфери имаш интерес да се развиваш и в кои по-слаб?"
column_titles[MENTOR_STUDIED] = "Какво си представяш, че е учил твоя ментор?"
column_titles[MENTOR_EXPERIENCE] = "Ментор в каква професионална сфера би бил/а най-полезен/а за теб?"
column_titles[SKILLS_TO_IMPROVE] = "Кои свои качества искаш да промениш/подобриш?"
column_titles[FREE_TIME_ACTIVITIES] = "Как се забавляваш в свободното си време?"
column_titles[DIFFICULT_SITUATION] = "Разкажи ни за трудна ситуация и как си се справил/а?"
column_titles[WHY_APPLY] = "Защо кандидатстваш в програмата?"
column_titles[IDEA_IN_ABLE_MENTOR] = "Каква идея искаш да осъществиш в рамките на ABLE Mentor?"
column_titles[WANT_TO_CHANGE] = "Желая да променя..."
column_titles[HOURS_PER_WEEK] = "По колко часа седмично би отделял/а на проекта?"
column_titles[PROJECT_WITH_MENTOR] = "По какъв проект би работил/а със своя ментор?"
column_titles[HEARD_OF_ABLE_MENTOR] = "Научил/а за ABLE Mentor от?"
column_titles[STUDENT_NAME_COPY] = "Ученик"
column_titles[MENTOR_NAME] = "Ментор"
column_titles[COORDINATOR_NAME] = "Координатор"


def try_create_doc(row_data, file_path):
    if row_data[STATUS] != "Активен":
        return False

    doc = docx.Document()
    doc.add_heading("ПРОФИЛ НА ТВОЯ УЧЕНИК", 0)

    table = doc.add_table(rows=0, cols=2)

    for idx in range(0, len(row_data)):
        if (idx == STATUS or
            idx == CONFIRMED or
            idx == STUDENT_NAME or
            idx == STUDENT_NAME_COPY or
            idx == MENTOR_NAME or
                idx not in column_titles):
            continue

        if idx > MENTOR_NAME:
            break

        table_row = table.add_row().cells
        table_row[0].text = column_titles[idx]
        table_row[1].text = row_data[idx]

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
            # coordinator_name = row[COORDINATOR_NAME].replace("/", "").strip()

            file_path = f"{OUTPUT_DIRECTORY}/{mentor_name}.docx".replace("\\", "/")
            # file_path = f"{OUTPUT_DIRECTORY}/({coordinator_name}) {mentor_name}.docx".replace("\\", "/")

            try_create_doc(row, file_path)


if __name__ == "__main__":
    create_docs()
