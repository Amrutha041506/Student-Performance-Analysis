"""
generate_dataset.py
-------------------
Creates the sample dataset  data/student_performance.csv

The data is random but reproducible (fixed seed), so you always get the
same file. A few mistakes are added ON PURPOSE (missing marks, duplicate
rows, impossible marks, invalid records) so the data-cleaning step in
analysis.py has real problems to find and fix - just like real-world data.

Run from the project folder:
    python src/generate_dataset.py
"""

from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------
# Settings
# ---------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent      # project folder
OUTPUT_FILE = BASE_DIR / "data" / "student_performance.csv"

NUMBER_OF_STUDENTS = 120
SUBJECTS = ["Mathematics", "Science", "English", "Social_Studies", "Computer"]

# Fixed seed = same dataset every time
rng = np.random.default_rng(42)

# ---------------------------------------------------------------
# Name lists used to build realistic student names
# ---------------------------------------------------------------
male_names = [
    "Aarav", "Vihaan", "Arjun", "Rohan", "Karthik", "Aditya", "Rahul", "Siddharth",
    "Nikhil", "Pranav", "Harsh", "Varun", "Kabir", "Manish", "Suresh", "Ravi",
    "Imran", "Faizan", "Daniel", "Joseph", "Amit", "Tarun", "Vivek", "Anil",
    "Sanjay", "Mohit", "Yash", "Dev", "Naveen", "Kiran",
]
female_names = [
    "Ananya", "Diya", "Priya", "Sneha", "Kavya", "Meera", "Isha", "Neha",
    "Pooja", "Riya", "Sana", "Aisha", "Fatima", "Lakshmi", "Swathi", "Divya",
    "Anjali", "Shruti", "Nandini", "Tanvi", "Bhavana", "Harini", "Sara", "Mary",
    "Pallavi", "Rashmi", "Aditi", "Keerthi", "Madhavi", "Zoya",
]
surnames = [
    "Sharma", "Reddy", "Patel", "Singh", "Kumar", "Nair", "Iyer", "Gupta",
    "Khan", "Rao", "Das", "Mehta", "Joshi", "Verma", "Naidu", "Menon",
    "Pillai", "Ali", "Shaikh", "Chowdhury", "Bose", "Kapoor", "Agarwal", "Yadav",
    "Fernandes", "Thomas", "Banerjee", "Goud", "Desai", "Chauhan",
]

# Small difference in difficulty for each subject (added to marks)
subject_effect = {
    "Mathematics": -3,
    "Science": 0,
    "English": 3,
    "Social_Studies": 2,
    "Computer": 4,
}

# ---------------------------------------------------------------
# Build the clean student records
# ---------------------------------------------------------------
records = []
used_names = set()

for i in range(NUMBER_OF_STUDENTS):
    gender = str(rng.choice(["Male", "Female"]))

    # Make sure every student gets a different full name
    while True:
        first_name = str(rng.choice(male_names if gender == "Male" else female_names))
        last_name = str(rng.choice(surnames))
        full_name = first_name + " " + last_name
        if full_name not in used_names:
            used_names.add(full_name)
            break

    student_class = int(rng.choice([9, 10, 11, 12]))
    age = student_class + 5 + int(rng.integers(0, 2))      # class 9 -> age 14 or 15

    # Each student has a general ability level; a few students are weaker
    if rng.random() < 0.06:
        ability = rng.normal(42, 8)
    else:
        ability = rng.normal(70, 11)

    row = {
        "Student_ID": "S" + str(i + 1).zfill(3),
        "Student_Name": full_name,
        "Gender": gender,
        "Age": age,
        "Class": student_class,
    }
    for subject in SUBJECTS:
        marks = ability + subject_effect[subject] + rng.normal(0, 9)
        row[subject] = int(np.clip(round(marks), 0, 100))
    records.append(row)

df = pd.DataFrame(records)

# Use nullable integer types so we can place missing values in the columns
df["Age"] = df["Age"].astype("Int64")
for subject in SUBJECTS:
    df[subject] = df[subject].astype("Int64")

# ---------------------------------------------------------------
# Add realistic data-quality problems (on purpose)
# ---------------------------------------------------------------
# 1) Missing marks (blank cells)
df.loc[12, "Mathematics"] = pd.NA
df.loc[88, "Mathematics"] = pd.NA
df.loc[45, "Science"] = pd.NA
df.loc[70, "English"] = pd.NA
df.loc[101, "English"] = pd.NA
df.loc[33, "Computer"] = pd.NA

# 2) A missing age
df.loc[41, "Age"] = pd.NA

# 3) Impossible marks (typing mistakes)
df.loc[19, "Mathematics"] = 105
df.loc[77, "English"] = -5
df.loc[96, "Computer"] = 150

# 4) Invalid age (impossible for a school student)
df.loc[66, "Age"] = 45

# 5) Text typed instead of a number ("AB" = absent)
df["Social_Studies"] = df["Social_Studies"].astype(object)
df.loc[54, "Social_Studies"] = "AB"

# 6) A student record without a name
df["Student_Name"] = df["Student_Name"].astype(object)
df.loc[9, "Student_Name"] = np.nan

# 7) Duplicate rows (same record entered twice)
duplicates = df.iloc[[3, 25, 60]]
df = pd.concat([df, duplicates], ignore_index=True)

# ---------------------------------------------------------------
# Save the CSV file
# ---------------------------------------------------------------
OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
df.to_csv(OUTPUT_FILE, index=False)

print("Dataset created:", OUTPUT_FILE)
print("Rows:", len(df), "| Columns:", len(df.columns))
