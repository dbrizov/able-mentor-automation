import io
import os
import csv
import docx
import matplotlib.pyplot as plt
import numpy


CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__))
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_self_evaluations"
RESPONSES_FILE_NAME = "self_evaluations_responses_beginning.csv"
RESPONSES_FILE_PATH = f"{CURRENT_DIRECTORY}/{RESPONSES_FILE_NAME}".replace("\\", "/")

texts = [None] * 37
texts[0] = "Timestamp"
texts[1] = "Email Address"

texts[2] = "Убедителност при говорене със събеседник"  # [Важност]
texts[3] = "Убедителност при говорене със събеседник"  # [Увереност]
texts[4] = "Умение за говорене пред публика"  # [Важност]
texts[5] = "Умение за говорене пред публика"  # [Увереност]
texts[6] = "Правилна онлайн комуникация по имейл"  # [Важност]
texts[7] = "Правилна онлайн комуникация по имейл"  # [Увереност]
texts[8] = "Получаване, даване на конструктивна обратна връзка"  # [Важност]
texts[9] = "Получаване, даване на конструктивна обратна връзка"  # [Увереност]
texts[10] = "Активно слушане"  # [Важност]
texts[11] = "Активно слушане"  # [Увереност]

texts[12] = "Успешна и ефективна работа в екип"  # [Важност]
texts[13] = "Успешна и ефективна работа в екип"  # [Увереност]
texts[14] = "Генериране на идеи"  # [Важност]
texts[15] = "Генериране на идеи"  # [Увереност]
texts[16] = "Структуриране на проект"  # [Важност]
texts[17] = "Структуриране на проект"  # [Увереност]
texts[18] = "Търсене и анализ на информация"  # [Важност]
texts[19] = "Търсене и анализ на информация"  # [Увереност]
texts[20] = "Изготвяне на презентация"  # [Важност]
texts[21] = "Изготвяне на презентация"  # [Увереност]
texts[22] = "Привличане на съмишленици, партньори, спонсори за реализиране на проект"  # [Важност]
texts[23] = "Привличане на съмишленици, партньори, спонсори за реализиране на проект"  # [Увереност]

texts[24] = "Поставяне на SMART цели"  # [Важност]
texts[25] = "Поставяне на SMART цели"  # [Увереност]
texts[26] = "Приоритизиране на задачи"  # [Важност]
texts[27] = "Приоритизиране на задачи"  # [Увереност]
texts[28] = "Ефективно планиране и управление на времето"  # [Важност]
texts[29] = "Ефективно планиране и управление на времето"  # [Увереност]
texts[30] = "Вътрешна мотивация"  # [Важност]
texts[31] = "Вътрешна мотивация"  # [Увереност]
texts[32] = "Ясна представа за следващите 5 години"  # [Важност]
texts[33] = "Ясна представа за следващите 5 години"  # [Увереност]

texts[34] = "Напиши най - важните 3 неща за теб, които искаш да постигнеш с участието си в програмата ABLE Mentor:"
texts[35] = "Казвам се..."
texts[36] = "Съгласен съм резултатът от моя тест да бъде даден за информация на моя ментор"


def create_bar_chart(title, bar_labels, data):
    x = numpy.arange(len(bar_labels))  # the label locations
    width = 0.2  # the width of the bars
    multiplier = 0

    fig, ax = plt.subplots(layout="constrained")
    fig.set_figwidth(7)
    fig.set_figheight(5.3)

    for attribute, measurement in data.items():
        offset = width * multiplier
        rects = ax.bar(x + offset, measurement, width, label=attribute)
        ax.bar_label(rects, padding=3)
        multiplier += 1

    ax.set_title(title)
    ax.set_xticks(x + width, bar_labels)
    ax.legend(loc="upper right", ncols=2)
    ax.set_ylim(0, 11.9)

    fig_bytes = io.BytesIO()
    plt.savefig(fig_bytes)
    plt.close()  # close current figure to free memory
    fig_bytes.seek(0)
    return fig_bytes


def try_create_doc(student_name, row_data, file_path):
    doc = docx.Document()

    # heading
    doc.add_heading(f"Резултат от тест за самооценка\n({student_name})", 0)

    # 3 important things
    p = doc.add_paragraph()
    p.add_run(f"{texts[34]}").bold = True
    doc.add_paragraph(f'"{row_data[34]}"')

    # communication
    bar_labels = [x.replace(" ", "\n") for x in texts[2:12:2]]
    data = {
        "Важност - начало": [int(x) for x in row_data[2:12:2]],
        "Увереност - начало": [int(x) for x in row_data[3:12:2]],
    }

    png = create_bar_chart("Комуникация", bar_labels, data)
    doc.add_picture(png)
    # plt.savefig(file_path.replace(".docx", ".png"))  # debug png

    # business skills
    bar_labels = [x.replace(" ", "\n") for x in texts[12:24:2]]
    data = {
        "Важност - начало": [int(x) for x in row_data[12:24:2]],
        "Увереност - начало": [int(x) for x in row_data[13:24:2]],
    }

    png = create_bar_chart("Бизнес умения", bar_labels, data)
    doc.add_picture(png)

    # personal effectiveness
    bar_labels = [x.replace(" ", "\n") for x in texts[24:34:2]]
    data = {
        "Важност - начало": [int(x) for x in row_data[24:34:2]],
        "Увереност - начало": [int(x) for x in row_data[25:34:2]],
    }

    png = create_bar_chart("Лична ефективност", bar_labels, data)
    doc.add_picture(png)

    doc.save(file_path)
    return True


def create_docs():
    if (not os.path.exists(OUTPUT_DIRECTORY)):
        os.mkdir(OUTPUT_DIRECTORY)

    with open(RESPONSES_FILE_PATH, encoding="utf-8", mode="r") as fstream:
        reader = csv.reader(fstream, delimiter=',', quotechar='"')

        for idx, row in enumerate(reader):
            if (idx == 0):
                continue  # skip first row

            student_name = row[35].replace("/", "").strip()
            file_path = f"{OUTPUT_DIRECTORY}/{student_name}.docx".replace("\\", "/")
            try_create_doc(student_name, row, file_path)


create_docs()
