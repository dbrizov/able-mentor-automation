import os
import docx
import pandas as pd


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles"
REGISTER_FILE_NAME = "register_students.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}"

# Only students with one of these statuses get a profile (matched case-insensitive)
PROFILE_STATUSES = ["in process", "matched"]

# Column names in the students' csv file (matched case-insensitive, long names by their beginning)
STUDENT_FIRST_NAME = "Име"
STUDENT_LAST_NAME = "Фамилия"
STUDENT_STATUS = "Статус"

# Rows of the profile table: (title, columns). The values of several columns are joined on separate lines.
PROFILE_FIELDS = [
    ("Удобен формат", ["Удобен Формат"]),
    ("В кой град предпочитате", ["В кой град предпочитате"]),
    ("Населено място", ["Населено Място"]),
    ("Възраст", ["Възраст"]),
    ("Ниво на английски език", ["Английски Език Ниво"]),
    ("Образование и кариера", ["Образование И Кариера"]),
    ("Училище", ["Училище"]),
    ("Завършен клас", ["Последно Завършен Клас"]),
    ("Любими предмети и дейности в училище", ["Любими Предмети И Дейности В Училище"]),
    ("Какво ще правя след гимназията", ["Какво Ще Правиш След Гимназията"]),
    ("За теб извън училище", ["За Теб Извън Училище"]),
    ("Интереси извън училище", ["Интереси Извън Училище"]),
    ("Разкажи ни за лично предизвикателство и как си го превъзмогнал/а?", ["Разкажете Ни За Ваше Лично Предизвикателство"]),
    ("Кой би бил ти в екип? (подредени по предпочитание)",
     ["Кой Би Бил Ти В Екип", "Избор 2", "Избор 3", "Избор 4", "Избор  5", "Избор 6"]),
    ("По колко часа седмично би отделял/а на проекта?", ["Средно по колко часа"]),
    ("Защо имаш нужда да участваш в програмата?", ["Защо имаш нужда да участваш"]),
    ("Участвал ли си в други сходни програми, завършил ли си ги и какво си взе от тях?",
     ["Участвал ли си в други сходни програми"]),
    ("По какъв проект би работил/а със своя ментор? (подредени по предпочитание)",
     ["По какъв проект", "проект 2", "проект 3", "проект 4"]),
    ("Можеш да добавиш няколко аспекта и да ги подредиш по важност.",
     ["Можеш да добавиш няколко аспекта и да ги подредиш по важност."]),
    ("Кое твърдение за теб по отношение на кариерната и академичната насока е вярно?",
     ["Кое твърдение за теб по отношение на кариерната"]),
    ("Кое твърдение за теб по отношение на участието ти в програмата е вярно?",
     ["Кое твърдение за теб по отношение на участието"]),
    ("Кажи повече за това твърдение", ["кажи повече за това твърдение"]),
    ("Сфери, които са ти интересни, и/или в които искаш да се развиеш",
     ["Сфери, които са ти интересни", "сфера 1", "повече за сфера 1", "сфера 2", "повече за сфера 2", "сфера 3",
      "повече за сфера 3"]),
    # pandas renames the 2nd of two equal headers to "<header>.1"
    ("Можеш да добавиш няколко аспекта и да ги подредиш по важност.",
     ["Можеш да добавиш няколко аспекта и да ги подредиш по важност..1"]),
    ("Има ли специфична сфера, различна от избраните, в която би искал/а менторът ти да има опит и защо?",
     ["Има ли специфична сфера"]),
    ("Кои аспекти на работата с ментора ще са най-полезни? (подредени по важност)",
     ["Кои аспекти на работата с ментора", "избери аспект 1", "избери аспект 2"]),
    ("Научил/а за ABLE Mentor от?", ["Откъде разбрахте"]),
]


def read_register(csv_file_path: str) -> pd.DataFrame:
    register = pd.read_csv(csv_file_path, dtype=str, encoding="utf-8").fillna("")
    register.columns = [column.strip() for column in register.columns]
    return register


def find_column(register: pd.DataFrame, column_name: str) -> str:
    column_name = column_name.casefold()
    for column in register.columns:
        if column.casefold() == column_name:
            return column

    columns = [column for column in register.columns if column.casefold().startswith(column_name)]
    if len(columns) != 1:
        raise KeyError(f"Expected one column starting with '{column_name}', found {len(columns)}")

    return columns[0]


def get_value(register: pd.DataFrame, row, column_name: str) -> str:
    return row[find_column(register, column_name)].strip()


def join_texts(texts: list[str]) -> str:
    # Several values are shown as a list, one "- value" per line
    texts = [text for text in texts if text and text != "---"]
    if len(texts) > 1:
        texts = [f"- {text}" for text in texts]

    return "\n".join(texts)


def create_doc(register: pd.DataFrame, row, file_path: str):
    doc = docx.Document()
    doc.add_heading("ПРОФИЛ НА ТВОЯ УЧЕНИК", 0)

    table = doc.add_table(rows=0, cols=2)

    for title, columns in PROFILE_FIELDS:
        value = join_texts([get_value(register, row, column) for column in columns])
        if not value:
            continue

        table_row = table.add_row().cells
        table_row[0].text = title
        table_row[1].text = value

    doc.save(file_path)


def create_docs():
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.mkdir(OUTPUT_DIRECTORY)

    register = read_register(REGISTER_FILE_PATH)
    students = list()
    for _, row in register.iterrows():
        name = f"{get_value(register, row, STUDENT_FIRST_NAME)} {get_value(register, row, STUDENT_LAST_NAME)}".strip()
        if not name:
            continue

        status = get_value(register, row, STUDENT_STATUS)
        if status.casefold() not in PROFILE_STATUSES:
            # print(f"skip student {name}: {status or 'no status'}")
            continue

        students.append((name, row))

    # Students with equal names get a number suffix, so their files don't overwrite each other
    name_counts = dict()
    for name, _ in students:
        name_counts[name] = name_counts.get(name, 0) + 1

    name_counters = dict()
    for name, row in students:
        file_name = name
        if name_counts[name] > 1:
            name_counters[name] = name_counters.get(name, 0) + 1
            file_name = f"{name}_{name_counters[name]}"

        file_path = f"{OUTPUT_DIRECTORY}/{file_name}.docx"
        create_doc(register, row, file_path)
        print(f"created {file_path}")


if __name__ == "__main__":
    create_docs()
