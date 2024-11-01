import os
import docx

CURRENT_DIRECTORY = os.path.dirname(
    os.path.realpath(__file__)).replace("\\", "/")
# The folder that stores the files (decide based on your need)
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles_redacted"

# Columns to remove (in the future you can change them just recplace with the title needed)
titles_to_remove = {
    "Ментор в каква професионална сфера би бил/а най-полезен/а за теб?",
    "Научил/а за ABLE Mentor от?"
}


def remove_specified_columns(doc_path):
    doc = docx.Document(doc_path)

    # Loop trough the tables to get the ones in need of modification
    for table in doc.tables:
        rows_to_delete = []
        for i, row in enumerate(table.rows):
            cell_text = row.cells[0].text.strip()
            if cell_text in titles_to_remove:
                rows_to_delete.append(i)

        # Delete the rows
        for row_idx in reversed(rows_to_delete):
            table._tbl.remove(table.rows[row_idx]._tr)

    doc.save(doc_path)


def process_all_documents():
    for file_name in os.listdir(OUTPUT_DIRECTORY):
        if file_name.endswith(".docx"):
            file_path = os.path.join(OUTPUT_DIRECTORY, file_name)
            remove_specified_columns(file_path)


if __name__ == "__main__":
    process_all_documents()