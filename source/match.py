import os
import re
import numpy as np
import pandas as pd
import xlsxwriter
from huggingface_hub import snapshot_download
from sentence_transformers import SentenceTransformer


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
STUDENTS_CSV_FILE_NAME = "register_students.csv"
STUDENTS_CSV_FILE_PATH = f"{CURRENT_DIRECTORY}/{STUDENTS_CSV_FILE_NAME}"
MENTORS_CSV_FILE_NAME = "register_mentors.csv"
MENTORS_CSV_FILE_PATH = f"{CURRENT_DIRECTORY}/{MENTORS_CSV_FILE_NAME}"
OUTPUT_FILE_NAME = "matches.xlsx"
OUTPUT_FILE_PATH = f"{CURRENT_DIRECTORY}/{OUTPUT_FILE_NAME}"
AI_MODEL_NAME = "BAAI/bge-m3"
AI_MODEL_FILE_PATH = f"{CURRENT_DIRECTORY}/ai_models/bge-m3"

# Number of mentor suggestions per student
TOP_N = 3

# Comparison weights
SEMANTIC_WEIGHT = 40
PROJECT_TYPE_WEIGHT = 25
AXES_WEIGHT = 20
FORMAT_WEIGHT = 10
HOURS_PER_WEEK_WEIGHT = 5

# Weights inside the semantic similarity
NEEDS_WEIGHT = 0.5  # what the student looks for vs what the mentor can give
SPHERES_WEIGHT = 0.5  # the student's spheres of interest vs the mentor's experience

# If True, a student who can only meet in person is never matched with an online-only mentor (and vice versa)
ENFORCE_FORMAT = True

# Column names in the students' csv file (matched case-insensitive, long names by their beginning)
STUDENT_FIRST_NAME = "Име"
STUDENT_LAST_NAME = "Фамилия"
STUDENT_EMAIL = "Email"
STUDENT_FORMAT = "Удобен Формат"
STUDENT_PROFILE = "Профил"
STUDENT_STATUS = "Статус"
STUDENT_MENTOR = "МЕНТОР"
STUDENT_PROPOSED_MENTOR = "Предложение за мачинг"
STUDENT_LOOKING_FOR = "търси се"  # needs
STUDENT_ADDITIONAL_INFO = "Допълнителна информация"  # needs
STUDENT_SPHERES = ["сфера 1", "повече за сфера 1", "сфера 2", "повече за сфера 2", "сфера 3", "повече за сфера 3"]  # spheres
STUDENT_MENTOR_SPHERE = "Има ли специфична сфера"  # spheres
STUDENT_PROJECT_TYPES = ["По какъв проект", "проект 2", "проект 3", "проект 4"]  # ordered by preference
STUDENT_HOURS_PER_WEEK = "Средно по колко часа"
STUDENT_STRUCTURE = "Нужда от структура"
STUDENT_SUPPORT = "Емоционална подкрепа"

# Column names in the mentors' csv file
MENTOR_FIRST_NAME = "Име"
MENTOR_LAST_NAME = "Фамилия"
MENTOR_EMAIL = "Email"
MENTOR_FORMAT = "Удобен Формат"
MENTOR_PROFILE = "Профил"
MENTOR_STATUS = "Статус"
MENTOR_SPHERES = ["Сфера 1", "Позиция 1", "Сфера 2", "Позиция 2", "Сфера 3", "Позиция 3"]  # spheres
MENTOR_HOBBIES = "Разкажете Ни За Вашите Интереси"  # spheres
MENTOR_PROJECT_TYPES = "По Какъв Проект"
MENTOR_HOURS_PER_WEEK = "Средно По Колко Часа"
MENTOR_STRUCTURE = "Нужда от структура"
MENTOR_SUPPORT = "Емоционална подкрепа"

# Sections of the generated profile
STUDENT_PROFILE_NEEDS = ["КАКВО ТЪРСИ ОТ МЕНТОР"]
STUDENT_PROFILE_SPHERES = ["СФЕРИ ЗА МЕНТОРА"]
MENTOR_PROFILE_NEEDS = ["КАКВО МОЖЕ ДА ДАДЕ НА УЧЕНИКА"]
MENTOR_PROFILE_SPHERES = ["СИЛНИ ТЕМИ", "ОПИТ", "ОБРАЗОВАНИЕ"]

# Project types are recognized by keyword, because the full labels differ between the two forms
PROJECT_TYPE_KEYWORDS = ["личностно", "бизнес", "практически", "социална"]
PROJECT_TYPE_CHOICE_SCORES = [1.0, 0.6, 0.4, 0.2]  # score for the student's 1st, 2nd, 3rd and 4th choice

# Formats
FORMAT_BOTH = "двата формата"
FORMAT_IN_PERSON = "на живо"
FORMAT_ONLINE = "онлайн"
FORMAT_PARTIAL_SCORE = 0.8  # one of the two accepts both formats

# Hours per week
MORE_HOURS_PER_WEEK = 5  # the "Повече" answer
MAX_HOURS_PER_WEEK_DIFF = 3

# Matching axes are in the range [1...5]
MAX_AXIS_DIFF = 4

# Score for a comparison where one of the values is missing
UNKNOWN_SCORE = 0.5


class PersonData:
    def __init__(self):
        self.name: str = ""
        self.email: str = ""
        self.format: str = ""
        self.needs: str = ""
        self.spheres: str = ""
        self.project_types: list[str] = list()
        self.hours_per_week: int | None = None
        self.structure: int | None = None
        self.support: int | None = None

    def __repr__(self):
        result = f"Name: {self.name}{os.linesep}"
        result += f"Format: {self.format}{os.linesep}"
        result += f"Needs: {self.needs}{os.linesep}"
        result += f"Spheres: {self.spheres}{os.linesep}"
        result += f"Project Types: {self.project_types}{os.linesep}"
        result += f"Hours/Week: {self.hours_per_week}{os.linesep}"
        result += f"Structure: {self.structure}{os.linesep}"
        result += f"Support: {self.support}{os.linesep}"
        return result


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


def parse_profile(profile: str) -> dict[str, str]:
    # A section starts with "■ HEADER" and continues until the next one; short sections are "■ HEADER: text"
    sections = dict()
    for section in re.split(r"^■ ", profile, flags=re.MULTILINE)[1:]:
        header, _, body = section.partition("\n")
        header, _, inline_body = header.partition(":")
        sections[header.strip()] = f"{inline_body.strip()}\n{body.strip()}".strip()

    return sections


def parse_format(format: str):
    format = format.replace("?", "").strip().lower()
    return format if format in (FORMAT_IN_PERSON, FORMAT_ONLINE) else FORMAT_BOTH


def parse_project_types(project_types: list[str]) -> list[str]:
    # Returns the project type keywords in the order they appear
    result = list()
    for project_type in project_types:
        found = [(project_type.lower().find(keyword), keyword) for keyword in PROJECT_TYPE_KEYWORDS]
        for index, keyword in sorted(found):
            if index >= 0 and keyword not in result:
                result.append(keyword)

    return result


def parse_hours_per_week(hours_per_week: str):
    match = re.search(r"\d+", hours_per_week)
    if match:
        return int(match.group())

    return MORE_HOURS_PER_WEEK if hours_per_week else None


def parse_axis(axis: str):
    try:
        return int(axis)
    except ValueError:
        return None


def join_texts(texts: list[str]) -> str:
    return "\n".join(text for text in texts if text and text != "---")


def extract_students(csv_file_path: str) -> list[PersonData]:
    register = read_register(csv_file_path)
    students = list()
    for _, row in register.iterrows():
        name = f"{get_value(register, row, STUDENT_FIRST_NAME)} {get_value(register, row, STUDENT_LAST_NAME)}".strip()
        if not name:
            continue

        # A status (in process, 50/50, matched...), a mentor or a proposed mentor means the student is already being matched
        status = (get_value(register, row, STUDENT_STATUS) or get_value(register, row, STUDENT_MENTOR) or
                  get_value(register, row, STUDENT_PROPOSED_MENTOR))
        if status:
            print(f"skip student {name}: {status}")
            continue

        profile = parse_profile(get_value(register, row, STUDENT_PROFILE))
        if not profile:
            print(f"skip student {name}: no profile")
            continue

        student = PersonData()
        student.name = name
        student.email = get_value(register, row, STUDENT_EMAIL)
        student.format = get_value(register, row, STUDENT_FORMAT)
        student.needs = join_texts([profile.get(section, "") for section in STUDENT_PROFILE_NEEDS] +
                                   [get_value(register, row, STUDENT_LOOKING_FOR),
                                    get_value(register, row, STUDENT_ADDITIONAL_INFO)])
        student.spheres = join_texts([profile.get(section, "") for section in STUDENT_PROFILE_SPHERES] +
                                     [get_value(register, row, column) for column in STUDENT_SPHERES] +
                                     [get_value(register, row, STUDENT_MENTOR_SPHERE)])
        student.project_types = parse_project_types([get_value(register, row, column) for column in STUDENT_PROJECT_TYPES])
        student.hours_per_week = parse_hours_per_week(get_value(register, row, STUDENT_HOURS_PER_WEEK))
        student.structure = parse_axis(get_value(register, row, STUDENT_STRUCTURE))
        student.support = parse_axis(get_value(register, row, STUDENT_SUPPORT))
        students.append(student)

    students.sort(key=lambda p: p.name)
    return students


def extract_mentors(csv_file_path: str) -> list[PersonData]:
    register = read_register(csv_file_path)
    mentors = list()
    for _, row in register.iterrows():
        name = f"{get_value(register, row, MENTOR_FIRST_NAME)} {get_value(register, row, MENTOR_LAST_NAME)}".strip()
        if not name:
            continue

        # Any status (in process, matched...) means the mentor is already being matched
        status = get_value(register, row, MENTOR_STATUS)
        if status:
            print(f"skip mentor {name}: {status}")
            continue

        profile = parse_profile(get_value(register, row, MENTOR_PROFILE))
        if not profile:
            print(f"skip mentor {name}: no profile")
            continue

        mentor = PersonData()
        mentor.name = name
        mentor.email = get_value(register, row, MENTOR_EMAIL)
        mentor.format = get_value(register, row, MENTOR_FORMAT)
        mentor.needs = join_texts([profile.get(section, "") for section in MENTOR_PROFILE_NEEDS])
        mentor.spheres = join_texts([profile.get(section, "") for section in MENTOR_PROFILE_SPHERES] +
                                    [get_value(register, row, column) for column in MENTOR_SPHERES] +
                                    [get_value(register, row, MENTOR_HOBBIES)])
        mentor.project_types = parse_project_types([get_value(register, row, MENTOR_PROJECT_TYPES)])
        mentor.hours_per_week = parse_hours_per_week(get_value(register, row, MENTOR_HOURS_PER_WEEK))
        mentor.structure = parse_axis(get_value(register, row, MENTOR_STRUCTURE))
        mentor.support = parse_axis(get_value(register, row, MENTOR_SUPPORT))
        mentors.append(mentor)

    mentors.sort(key=lambda p: p.name)
    return mentors


def load_model() -> SentenceTransformer:
    if not os.path.isfile(f"{AI_MODEL_FILE_PATH}/modules.json"):
        print(f"download {AI_MODEL_NAME} to {AI_MODEL_FILE_PATH}")
        snapshot_download(AI_MODEL_NAME, local_dir=AI_MODEL_FILE_PATH,
                          ignore_patterns=["onnx/*", "imgs/*", "*.jpg", ".gitattributes"])

    return SentenceTransformer(AI_MODEL_FILE_PATH)


def get_semantic_sim(students: list[PersonData], mentors: list[PersonData]) -> np.ndarray:
    model = load_model()

    def encode(texts):
        return model.encode(texts, batch_size=8, normalize_embeddings=True, show_progress_bar=True)

    # The embeddings are normalized, so the dot product is the cosine similarity
    needs_sim = encode([s.needs for s in students]) @ encode([m.needs for m in mentors]).T
    spheres_sim = encode([s.spheres for s in students]) @ encode([m.spheres for m in mentors]).T
    sim = NEEDS_WEIGHT * needs_sim + SPHERES_WEIGHT * spheres_sim

    # Mentors with a general profile are similar to every student, so each mentor's average similarity is subtracted.
    # What remains is how much better the mentor fits this student than a typical student.
    sim = sim - sim.mean(axis=0, keepdims=True)

    # The cosine similarities are in a narrow range, so they are rescaled per student to [0...1],
    # otherwise the other comparisons would outweigh them
    min_sim = sim.min(axis=1, keepdims=True)
    max_sim = sim.max(axis=1, keepdims=True)
    return (sim - min_sim) / np.maximum(max_sim - min_sim, 1e-9)


def get_project_type_sim(student: PersonData, mentor: PersonData) -> float:
    for project_type, score in zip(student.project_types, PROJECT_TYPE_CHOICE_SCORES):
        if project_type in mentor.project_types:
            return score

    return 0.0


def get_axes_sim(student: PersonData, mentor: PersonData) -> float:
    if student.structure is None or student.support is None or mentor.structure is None or mentor.support is None:
        return UNKNOWN_SCORE

    diff = abs(student.structure - mentor.structure) + abs(student.support - mentor.support)
    return 1 - diff / (2 * MAX_AXIS_DIFF)


def get_format_sim(student: PersonData, mentor: PersonData) -> float:
    student_format = parse_format(student.format)
    mentor_format = parse_format(mentor.format)
    if student_format == mentor_format:
        return 1.0

    if FORMAT_BOTH in (student_format, mentor_format):
        return FORMAT_PARTIAL_SCORE

    return 0.0


def get_hours_per_week_sim(student: PersonData, mentor: PersonData) -> float:
    if student.hours_per_week is None or mentor.hours_per_week is None:
        return UNKNOWN_SCORE

    return max(0.0, 1 - abs(student.hours_per_week - mentor.hours_per_week) / MAX_HOURS_PER_WEEK_DIFF)


def to_percent(value: float) -> int:
    return int(round(value * 100))


def save_matches(rows: list[list]):
    # One row per suggestion, so the sheet can be sorted and filtered
    headers = ["Student", "Student Email", "Student Format", "Rank", "Mentor", "Mentor Email", "Mentor Format",
               "Score", "Semantic", "Project Type", "Axes", "Format", "Hours/Week"]
    column_widths = [28, 32, 16, 6, 28, 32, 16, 8, 10, 13, 8, 8, 12]

    workbook = xlsxwriter.Workbook(OUTPUT_FILE_PATH)
    worksheet = workbook.add_worksheet("Matches")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "border": 1})
    student_formats = [workbook.add_format(), workbook.add_format({"bg_color": "#F2F2F2"})]  # alternate per student

    worksheet.write_row(0, 0, headers, header_format)
    for col_i, column_width in enumerate(column_widths):
        worksheet.set_column(col_i, col_i, column_width)

    student_i = -1
    for row_i, row in enumerate(rows, start=1):
        if row_i == 1 or row[0] != rows[row_i - 2][0]:
            student_i += 1
        worksheet.write_row(row_i, 0, row, student_formats[student_i % 2])

    worksheet.freeze_panes(1, 1)
    worksheet.autofilter(0, 0, len(rows), len(headers) - 1)
    workbook.close()


def find_matches():
    students = extract_students(STUDENTS_CSV_FILE_PATH)
    mentors = extract_mentors(MENTORS_CSV_FILE_PATH)
    print(f"students: {len(students)}, mentors: {len(mentors)}")

    semantic_sim = get_semantic_sim(students, mentors)
    max_sim = SEMANTIC_WEIGHT + PROJECT_TYPE_WEIGHT + AXES_WEIGHT + FORMAT_WEIGHT + HOURS_PER_WEEK_WEIGHT

    print("find matches:")
    rows = list()
    for student_i, student in enumerate(students):
        mentor_suggestions = list()
        for mentor_i, mentor in enumerate(mentors):
            sims = {
                "semantic": float(semantic_sim[student_i, mentor_i]),
                "project": get_project_type_sim(student, mentor),
                "axes": get_axes_sim(student, mentor),
                "format": get_format_sim(student, mentor),
                "hours": get_hours_per_week_sim(student, mentor),
            }
            if ENFORCE_FORMAT and sims["format"] == 0:
                continue

            final_sim = (sims["semantic"] * SEMANTIC_WEIGHT + sims["project"] * PROJECT_TYPE_WEIGHT +
                         sims["axes"] * AXES_WEIGHT + sims["format"] * FORMAT_WEIGHT +
                         sims["hours"] * HOURS_PER_WEEK_WEIGHT)
            mentor_suggestions.append((mentor_i, final_sim / max_sim, sims))

        mentor_suggestions.sort(key=lambda x: x[1], reverse=True)

        for rank, (mentor_i, sim, sims) in enumerate(mentor_suggestions[:TOP_N], start=1):
            mentor = mentors[mentor_i]
            rows.append([student.name, student.email, student.format, rank, mentor.name, mentor.email,
                         mentor.format, to_percent(sim), *[to_percent(sim_value) for sim_value in sims.values()]])
            print(f"{student.name}(Y) + {mentor.name}(M) - {to_percent(sim)}")

    save_matches(rows)
    print(f"saved {OUTPUT_FILE_PATH}")


if __name__ == "__main__":
    find_matches()
