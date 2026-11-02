import os
import re
import html
import logging
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
PART_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_part.html"
CARD_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_card.html"
FIELD_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_field.html"
RATING_TEMPLATE_FILE_PATH = f"{HTML_DIRECTORY}/student_profile_rating.html"

logging.basicConfig(filename=f"{CURRENT_DIRECTORY}/student_profiles.log", filemode="w", encoding="utf-8",
                    level=logging.INFO, format="%(message)s")

APPLICATION_PART_TITLE = "Кандидатура на ученика"

# Theme of the cards, a CSS class in css/student_profile.css
WARM_THEME = "warm"

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
STUDENT_MENTOR = "Предложение за ментор"  # the profile's file is named <student>_<mentor>

# Kinds of fields: a text, a rating from 1 to 5 (shown as dots), or choices separated by commas (shown as a list)
TEXT_FIELD = "text"
RATING_FIELD = "rating"
CHOICES_FIELD = "choices"

# Sections of the profile: (section title, fields). Each field is a title and a value in the section's card:
# (title, columns) or (title, columns, kind). The values of several columns are joined on separate lines.
# Empty fields and sections are skipped.
APPLICATION_SECTIONS = [
    ("Основна информация", [
        ("Удобен формат", ["Удобен Формат"]),
        ("В кой град предпочитате", ["В кой град предпочитате"]),
        ("Населено място", ["Населено Място"]),
        ("Възраст", ["Възраст"]),
        ("Ниво на английски език", ["Английски Език Ниво"]),
    ]),
    ("Образование и кариера", [
        ("Училище", ["Училище"]),
        ("Завършен клас", ["Последно Завършен Клас"]),
        ("Любими предмети и дейности в училище", ["Любими Предмети И Дейности В Училище"]),
        ("Какво ще правя след гимназията", ["Какво Ще Правиш След Гимназията"]),
    ]),
    ("За теб извън училище", [
        ("Интереси извън училище", ["Интереси Извън Училище"]),
        ("Разкажи ни за лично предизвикателство и как си го превъзмогнал/а?",
         ["Разкажете Ни За Ваше Лично Предизвикателство"]),
        ("Кой би бил ти в екип? (подредени по предпочитание)",
         ["Кой Би Бил Ти В Екип", "Избор 2", "Избор 3", "Избор 4", "Избор  5"]),
    ]),
    ("Участие в програмата", [
        ("По колко часа седмично би отделял/а на проекта?", ["Средно по колко часа"]),
        ("Защо имаш нужда да участваш в програмата?", ["Защо имаш нужда да участваш"]),
        ("Участвал ли си в други сходни програми, завършил ли си ги и какво си взе от тях?",
         ["Участвал ли си в други сходни програми"]),
        ("По какъв проект би работил/а със своя ментор? (подредени по предпочитание)",
         ["По какъв проект", "проект 2", "проект 3", "проект 4"]),
        ("Кое твърдение за теб по отношение на кариерната и академичната насока е вярно?",
         ["Кое твърдение за теб по отношение на кариерната"]),
        ("Кое твърдение за теб по отношение на участието ти в програмата е вярно?",
         ["Кое твърдение за теб по отношение на участието"]),
    ]),
    ("Сфери и ментор", [
        ("Сфери, които са ти интересни, и/или в които искаш да се развиеш",
         ["сфера 1", "повече за сфера 1", "сфера 2", "повече за сфера 2", "сфера 3",
          "повече за сфера 3"]),
        ("Има ли специфична сфера, различна от избраните, в която би искал/а менторът ти да има опит и защо?",
         ["Има ли специфична сфера"]),
        ("Кои аспекти на работата с ментора ще са най-полезни? (подредени по важност)",
         ["Кои аспекти на работата с ментора", "избери аспект 1", "избери аспект 2"]),
        ("Научил/а за ABLE Mentor от?", ["Откъде разбрахте"]),
    ]),
]


def log(message: str):
    print(message)
    logging.info(message)


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


def get_rows(register: pd.DataFrame, row, fields: list[tuple]) -> list[tuple[str, str, str]]:
    # The (title, value, kind) of the fields with a value
    rows = list()
    for title, columns, *kind in fields:
        kind = kind[0] if kind else TEXT_FIELD
        texts = [get_value(register, row, column) for column in columns]
        if kind == CHOICES_FIELD:
            # The choices are separated by ", " and start with a capital letter
            texts = [choice for text in texts for choice in re.split(r",\s*(?=[А-ЯA-Z])", text)]

        value = join_texts(texts)
        if value and title not in HIDDEN_FIELDS:
            rows.append((title, value, kind))

    return rows


# ---------- HTML ----------

def to_url(file_path: str) -> str:
    return Path(file_path).resolve().as_uri()


def read_template(file_path: str) -> Template:
    # An HTML template with $placeholders
    with open(file_path, encoding="utf-8", mode="r") as fstream:
        return Template(fstream.read())


def make_field(title: str, value: str, kind: str) -> str:
    if kind == RATING_FIELD and value in ["1", "2", "3", "4", "5"]:
        lines = [read_template(RATING_TEMPLATE_FILE_PATH).substitute(value=value)]
    else:
        # Every line of the value is a div, the "- " lines are bullets with a hanging indent
        lines = [f'<div class="bullet">{html.escape(line)}</div>' if line.startswith("- ")
                 else f"<div>{html.escape(line)}</div>" for line in value.split("\n")]

    return read_template(FIELD_TEMPLATE_FILE_PATH).substitute(title=html.escape(title), lines="\n".join(lines))


def make_card(number: int, title: str, rows: list[tuple[str, str, str]], theme: str) -> str:
    fields = [make_field(field_title, value, kind) for field_title, value, kind in rows]
    return read_template(CARD_TEMPLATE_FILE_PATH).substitute(
        number=f"{number:02d}", title=html.escape(title), theme=theme, fields="\n".join(fields))


def make_part(title: str, table: pd.DataFrame, row, sections: list[tuple], theme: str) -> str:
    # A part of the profile with a card for every section, the cards are numbered from 1 in every part
    cards = list()
    for section_title, fields in sections:
        rows = get_rows(table, row, fields)
        if rows:
            cards.append(make_card(len(cards) + 1, section_title, rows, theme))

    return read_template(PART_TEMPLATE_FILE_PATH).substitute(
        title=html.escape(title), theme=theme, cards="\n".join(cards))


def make_html(register: pd.DataFrame, row) -> str:
    part = make_part(APPLICATION_PART_TITLE, register, row, APPLICATION_SECTIONS, WARM_THEME)
    return read_template(PAGE_TEMPLATE_FILE_PATH).substitute(base_url=f"{to_url(HTML_DIRECTORY)}/", parts=part)


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
    profiles = list()  # (file name, row)
    for _, row in register.iterrows():
        student_name = (f"{get_value(register, row, STUDENT_FIRST_NAME)} "
                        f"{get_value(register, row, STUDENT_LAST_NAME)}").strip()
        if not student_name:
            continue

        status = get_value(register, row, STUDENT_STATUS)
        if status.casefold() not in PROFILE_STATUSES:
            print(f"skip student {student_name}: {status or 'no status'}")
            continue

        mentor_name = get_value(register, row, STUDENT_MENTOR)
        if not mentor_name:
            log(f"no mentor for {student_name}, the profile is named after the student only")

        profiles.append((f"{student_name}_{mentor_name}" if mentor_name else student_name, row))

    # Equal file names get a number suffix, so the files don't overwrite each other
    name_counts = dict()
    for name, _ in profiles:
        name_counts[name] = name_counts.get(name, 0) + 1

    for name, count in name_counts.items():
        if count > 1:
            log(f"{count} profiles are named {name}, they get a number suffix")

    # The profiles are HTML pages, printed to PDF by Playwright's Chromium (installed with "playwright install chromium")
    name_counters = dict()
    with tempfile.TemporaryDirectory() as temp_directory, sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page()
        for name, row in profiles:
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
