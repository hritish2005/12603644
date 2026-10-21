from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime, timedelta
import subprocess

# --------------------------------------------------
# SETTINGS
# --------------------------------------------------

REPO = Path(__file__).resolve().parent
EXCEL_FILE = REPO / "12603644.xlsx"
SHEET_NAME = "Daily Log"
LOG_FOLDER = REPO / "logs"


# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------

def ask_number(question):
    while True:
        try:
            return int(input(question).strip())
        except ValueError:
            print("Please enter a number.")


def ask_choice(question, choices):
    print()
    print(question)

    for i, choice in enumerate(choices, 1):
        print(f"  {i}. {choice}")

    while True:
        try:
            number = int(input("Choose: "))

            if 1 <= number <= len(choices):
                return choices[number - 1]

        except ValueError:
            pass

        print("Please choose a valid option.")


# --------------------------------------------------
# CHECK FILE
# --------------------------------------------------

if not EXCEL_FILE.exists():
    print("\nERROR: Excel file not found:")
    print(EXCEL_FILE)
    input("\nPress Enter to exit...")
    raise SystemExit


# --------------------------------------------------
# LOAD EXCEL
# --------------------------------------------------

workbook = load_workbook(EXCEL_FILE)
sheet = workbook[SHEET_NAME]

today = datetime.now().date()

last_date = None
last_row = None

# --------------------------------------------------
# FIND LAST LOGGED DATE
# --------------------------------------------------

for row in range(6, sheet.max_row + 1):

    value = sheet.cell(row, 1).value

    if isinstance(value, datetime):
        current_date = value.date()

    elif value:
        try:
            current_date = datetime.strptime(
                str(value),
                "%Y-%m-%d"
            ).date()

        except ValueError:
            continue

    else:
        continue

    if last_date is None or current_date > last_date:
        last_date = current_date
        last_row = row


# --------------------------------------------------
# DETERMINE NEXT DATE
# --------------------------------------------------

if last_date:
    next_date = last_date + timedelta(days=1)
else:
    next_date = today


# --------------------------------------------------
# DUPLICATE PROTECTION
# --------------------------------------------------

if last_date == today:

    print("\n========================================")
    print("          ALREADY LOGGED TODAY")
    print("========================================")

    print(
        f"\nYou already have a log for "
        f"{today.strftime('%d %B %Y')}."
    )

    print("\nNo changes were made.")

    input("\nPress Enter to exit...")
    raise SystemExit


# --------------------------------------------------
# SHOW LOGGER
# --------------------------------------------------

print("\n========================================")
print("          DAILY TASK LOGGER")
print("========================================")

print(
    f"\nNew entry: {next_date.strftime('%d %B %Y')}"
)

print("\nDurations are in MINUTES.\n")


# --------------------------------------------------
# COLLECT DATA
# --------------------------------------------------

sleep = ask_number("Sleep (minutes): ")

fitness = ask_number("Fitness (minutes): ")

study = ask_number("Study (minutes): ")

coding = ask_number("Coding (minutes): ")

class_time = ask_number("Class (minutes): ")

classes_attended = ask_number(
    "Classes attended: "
)

other = ask_number(
    "Other activities (minutes): "
)


feeling = ask_choice(
    "Today's feeling:",
    [
        "Low",
        "Neutral",
        "Good"
    ]
)


satisfaction = ask_choice(
    "Satisfaction level:",
    [
        "Unsatisfied",
        "Neutral",
        "Satisfied"
    ]
)


energy = ask_choice(
    "Energy level:",
    [
        "Low",
        "Medium",
        "High"
    ]
)


notes = input("\nNotes: ").strip()


# --------------------------------------------------
# FIND EMPTY EXCEL ROW
# --------------------------------------------------

target_row = None

for row in range(6, sheet.max_row + 1):

    if sheet.cell(row, 1).value is None:

        target_row = row
        break


if target_row is None:
    target_row = sheet.max_row + 1


# --------------------------------------------------
# WRITE DATA
# --------------------------------------------------

sheet.cell(target_row, 1).value = next_date
sheet.cell(target_row, 2).value = sleep
sheet.cell(target_row, 3).value = fitness
sheet.cell(target_row, 4).value = study
sheet.cell(target_row, 5).value = coding
sheet.cell(target_row, 6).value = class_time
sheet.cell(target_row, 7).value = classes_attended
sheet.cell(target_row, 8).value = other


# Total Tracked
sheet.cell(target_row, 9).value = (
    f'=IF(SUM(B{target_row}:F{target_row},'
    f'H{target_row})=0,"",'
    f'SUM(B{target_row}:F{target_row},'
    f'H{target_row}))'
)


# Free / Unaccounted
sheet.cell(target_row, 10).value = (
    f'=IF(I{target_row}="","",'
    f'1440-I{target_row})'
)


sheet.cell(target_row, 11).value = feeling
sheet.cell(target_row, 12).value = satisfaction
sheet.cell(target_row, 13).value = energy
sheet.cell(target_row, 14).value = notes


# --------------------------------------------------
# COPY FORMATTING
# --------------------------------------------------

if target_row > 6:

    source_row = target_row - 1

    for column in range(1, 15):

        source = sheet.cell(
            source_row,
            column
        )

        target = sheet.cell(
            target_row,
            column
        )

        if source.has_style:
            target._style = source._style

        target.number_format = source.number_format


sheet.cell(
    target_row,
    1
).number_format = "dd-mmm-yyyy"


# --------------------------------------------------
# SAVE EXCEL
# --------------------------------------------------

workbook.save(EXCEL_FILE)


# --------------------------------------------------
# CREATE DAILY MARKDOWN LOG
# --------------------------------------------------

LOG_FOLDER.mkdir(exist_ok=True)

date_string = next_date.strftime("%Y-%m-%d")

log_file = LOG_FOLDER / f"{date_string}.md"


total_tracked = (
    sleep
    + fitness
    + study
    + coding
    + class_time
    + other
)

free_time = 1440 - total_tracked


log_content = f"""# Daily Log — {next_date.strftime("%d %B %Y")}

## Time

| Activity | Minutes |
|---|---:|
| Sleep | {sleep} |
| Fitness | {fitness} |
| Study | {study} |
| Coding | {coding} |
| Class | {class_time} |
| Other Activities | {other} |
| **Total Tracked** | **{total_tracked}** |
| **Free / Unaccounted** | **{free_time}** |

## Academic

- Classes attended: **{classes_attended}**

## Reflection

- **Feeling:** {feeling}
- **Satisfaction:** {satisfaction}
- **Energy:** {energy}

## Notes

{notes if notes else "_No notes._"}
"""


log_file.write_text(
    log_content,
    encoding="utf-8"
)


# --------------------------------------------------
# GIT HELPER
# --------------------------------------------------

def git(command):

    result = subprocess.run(
        command,
        cwd=REPO,
        text=True,
        capture_output=True
    )

    if result.returncode != 0:

        print("\nGit error:")
        print(result.stderr)

        raise SystemExit

    return result.stdout.strip()


# --------------------------------------------------
# GIT ADD
# --------------------------------------------------

print("\n----------------------------------------")
print("Updating GitHub...")
print("----------------------------------------")


git([
    "git",
    "add",
    "12603644.xlsx",
    "logs/"
])


# --------------------------------------------------
# GIT COMMIT
# --------------------------------------------------

commit_message = (
    f"Daily log - {date_string}"
)


commit = subprocess.run(
    [
        "git",
        "commit",
        "-m",
        commit_message
    ],
    cwd=REPO,
    text=True,
    capture_output=True
)


if commit.returncode != 0:

    print("\nGit commit failed:")
    print(commit.stdout)
    print(commit.stderr)

    input("\nPress Enter to exit...")
    raise SystemExit


# --------------------------------------------------
# GIT PUSH
# --------------------------------------------------

push = subprocess.run(
    [
        "git",
        "push"
    ],
    cwd=REPO,
    text=True,
    capture_output=True
)


if push.returncode != 0:

    print("\nGit push failed:")
    print(push.stdout)
    print(push.stderr)

    input("\nPress Enter to exit...")
    raise SystemExit


# --------------------------------------------------
# SUCCESS
# --------------------------------------------------

print("\n========================================")
print("              SUCCESS")
print("========================================")

print(
    f"\nDate: {next_date.strftime('%d %B %Y')}"
)

print(
    f"Excel: {EXCEL_FILE.name}"
)

print(
    f"Log: logs/{date_string}.md"
)

print(
    f"Commit: {commit_message}"
)

print("\nGitHub has been updated.")

input("\nPress Enter to exit...")