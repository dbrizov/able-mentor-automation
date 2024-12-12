import os 
import csv
import math

CURRENT_DIRECTORY = os.path.dirname(
    os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles_redacted"
REGISTER_FILE_NAME = "student_register.csv"
REGISTER_FILE_PATH = f"{CURRENT_DIRECTORY}/{REGISTER_FILE_NAME}"


def get_column_index(column):
    decimal_value = 0
    for idx in reversed(range(0, len(column))):
        decimal_value += (ord(column[idx]) - ord("A") + 1) * \
            math.pow(26, len(column) - idx - 1)
    return int(decimal_value - 1)


STUDENT_NAME = get_column_index("D")
MENTOR_NAME = get_column_index("CC")


def rename_existing_profiles():
    mentor_student_map = {}
    try:
        with open(REGISTER_FILE_PATH, encoding="utf-8", mode="r") as fstream:
            reader = csv.reader(fstream, delimiter=',', quotechar='"')
            next(reader)  # Skip header row

            for row in reader:
                student_name = row[STUDENT_NAME].strip()
                mentor_name = row[MENTOR_NAME].strip()
                if mentor_name:
                    mentor_student_map[student_name] = mentor_name
    except Exception as e:
        print(f"Error reading the CSV file: {e}")
        return

    for filename in os.listdir(OUTPUT_DIRECTORY):
        if filename.endswith(".docx"):
            try:
                parts = filename.split("_")
                if len(parts) < 2:
                    print(f"Invalid filename format: '{filename}'")
                    continue

                student_name_in_file = parts[1].replace(".docx", "").strip()

                matched_student_name = None
                for student_name in mentor_student_map.keys():
                    if student_name_in_file == student_name or \
                       student_name_in_file == f"{student_name}1" or \
                       student_name_in_file == f"{student_name}B":
                        matched_student_name = student_name
                        break

                if matched_student_name:
                    mentor_name = mentor_student_map[matched_student_name]
                    new_filename = f"{mentor_name}.docx"
                    os.rename(os.path.join(OUTPUT_DIRECTORY, filename),
                              os.path.join(OUTPUT_DIRECTORY, new_filename))
                    print(f"Renamed '{filename}' to '{new_filename}'")
                else:
                    print(
                        f"No mentor found for student: {student_name_in_file}")
            except IndexError as e:
                print(f"Error processing file '{filename}': {e}")
            except Exception as e:
                print(
                    f"An error occurred while renaming file '{filename}': {e}")


if __name__ == "__main__":
    rename_existing_profiles()