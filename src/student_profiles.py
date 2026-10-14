import os
import html
import tempfile
import pandas as pd
from pathlib import Path
from string import Template
from playwright.sync_api import Page, sync_playwright


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles"
REGISTER_FILE_NAME = "register_students.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}"
HTML_DIRECTORY = f"{os.path.dirname(CURRENT_DIRECTORY)}/html"
PAGE_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_page.html"
CARD_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_card.html"
FIELD_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_field.html"

# Themes of the cards, CSS classes in css/student_profile.css.
# The student's data is on warm cards, the information from the team is on cold ones.
WARM_THEME = "warm"
COLD_THEME = "cold"

# Fields that are not shown in the profiles (by their title)
HIDDEN_FIELDS = [
    "Научил/а за ABLE Mentor от?",
]

# Only students with one of these statuses get a profile (matched case-insensitive)
PROFILE_STATUSES = ["in process", "matched"]

# Column names in the students' csv file (matched case-insensitive, long names by their beginning)
STUDENT_FIRST_NAME = "Име"
STUDENT_LAST_NAME = "Фамилия"
STUDENT_STATUS = "Статус"

# Sections of the profile: (section title, fields). Each field is a title and a value in the section's card:
# (title, columns). The values of several columns are joined on separate lines. Empty fields and sections are skipped.
PROFILE_SECTIONS = [
    ("Основна информация", [
        ("Удобен формат", ["Удобен Формат"]),
        ("В кой град предпочитате", ["В кой град предпочитате"]),
        ("Населено място", ["Населено Място"]),
        ("Възраст", ["Възраст"]),
        ("Ниво на английски език", ["Английски Език Ниво"]),
    ]),
    ("Образование и кариера", [
        ("Образование и кариера", ["Образование И Кариера"]),
        ("Училище", ["Училище"]),
        ("Завършен клас", ["Последно Завършен Клас"]),
        ("Любими предмети и дейности в училище", ["Любими Предмети И Дейности В Училище"]),
        ("Какво ще правя след гимназията", ["Какво Ще Правиш След Гимназията"]),
    ]),
    ("За теб извън училище", [
        ("За теб извън училище", ["За Теб Извън Училище"]),
        ("Интереси извън училище", ["Интереси Извън Училище"]),
        ("Разкажи ни за лично предизвикателство и как си го превъзмогнал/а?",
         ["Разкажете Ни За Ваше Лично Предизвикателство"]),
        ("Кой би бил ти в екип? (подредени по предпочитание)",
         ["Кой Би Бил Ти В Екип", "Избор 2", "Избор 3", "Избор 4", "Избор  5", "Избор 6"]),
    ]),
    ("Участие в програмата", [
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
    ]),
    ("Сфери и ментор", [
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
    ]),
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


def get_rows(register: pd.DataFrame, row, fields: list[tuple[str, list[str]]]) -> list[tuple[str, str]]:
    # The (title, value) of the fields with a value
    rows = list()
    for title, columns in fields:
        value = join_texts([get_value(register, row, column) for column in columns])
        if value and title not in HIDDEN_FIELDS:
            rows.append((title, value))

    return rows


# ---------- HTML ----------

def to_url(file_path: str) -> str:
    return Path(file_path).resolve().as_uri()


def read_template(file_path: str) -> Template:
    # An HTML template with $placeholders
    with open(file_path, encoding="utf-8", mode="r") as fstream:
        return Template(fstream.read())


def make_field(title: str, value: str) -> str:
    # Every line of the value is a div, the "- " lines are bullets with a hanging indent
    lines = [f'<div class="bullet">{html.escape(line)}</div>' if line.startswith("- ")
             else f"<div>{html.escape(line)}</div>" for line in value.split("\n")]
    return read_template(FIELD_TEMPLATE_FILE_PATH).substitute(title=html.escape(title), lines="\n".join(lines))


def make_card(number: int, title: str, rows: list[tuple[str, str]], theme: str) -> str:
    fields = [make_field(field_title, value) for field_title, value in rows]
    return read_template(CARD_TEMPLATE_FILE_PATH).substitute(
        number=f"{number:02d}", title=html.escape(title), theme=theme, fields="\n".join(fields))


def make_html(register: pd.DataFrame, row) -> str:
    cards = list()
    for section_title, fields in PROFILE_SECTIONS:
        rows = get_rows(register, row, fields)
        if rows:
            cards.append(make_card(len(cards) + 1, section_title, rows, WARM_THEME))

    return read_template(PAGE_TEMPLATE_FILE_PATH).substitute(base_url=f"{to_url(HTML_DIRECTORY)}/",
                                                             cards="\n".join(cards))


# ---------- PDF ----------

def create_pdf(register: pd.DataFrame, row, file_path: str, page: Page, temp_directory: str):
    # The HTML is a file, so it can use the local pictures and font. The PDF has the page size and margins of the CSS.
    html_file_path = f"{temp_directory}/profile.html"
    with open(html_file_path, encoding="utf-8", mode="w") as fstream:
        fstream.write(make_html(register, row))

    page.goto(to_url(html_file_path))
    page.evaluate("document.fonts.ready")
    page.pdf(path=file_path, print_background=True, prefer_css_page_size=True)


def create_pdfs():
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

    # The profiles are HTML pages, printed to PDF by Playwright's Chromium (installed with "playwright install chromium")
    name_counters = dict()
    with tempfile.TemporaryDirectory() as temp_directory, sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page()
        for name, row in students:
            file_name = name
            if name_counts[name] > 1:
                name_counters[name] = name_counters.get(name, 0) + 1
                file_name = f"{name}_{name_counters[name]}"

            file_path = f"{OUTPUT_DIRECTORY}/{file_name}.pdf"
            create_pdf(register, row, file_path, page, temp_directory)
            print(f"created {file_path}")

        browser.close()


if __name__ == "__main__":
    create_pdfs()
