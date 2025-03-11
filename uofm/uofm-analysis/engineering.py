# to run
# python3 engineering.py > analysis.txt

# this will give you
# 1.  analysis in a text file : analysis.txt
# 2.  flagged students in an excel file : flagged_students.xlsx
# 3.  cleaned data for analysis in an excel file : output_analysis.xlsx

import subprocess
import sys
import pandas as pd

# Ensure pandas is installed
try:
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
    import pandas as pd

# Engineering programs AE
programs = ["03-BE", "03-TE", "03-M", "03-C", "03-E"]

# Started with 202190 because 202090 is their first DE term
years = [202190, 202290, 202390, 202490]

# Read the Excel file
df = pd.read_excel("eng_analysis.xlsx")

# Create separate DataFrames for DE and AE records
de = df[df["Applicant Type"].str.lower() == "direct entry"]

ae = df[
    (df["Applicant Type"].str.lower() == "advanced entry") &
    (df["AE App prog is first choice"].str.lower() == "yes") &
    (~df["Current Engineering Program"].str.contains("ENGAP \\(Engineering Access Program\\)", case=False, na=False))
]

# Merge DE and AE on student ID
merged = pd.merge(ae, de, on="Student ID", suffixes=("_ae", "_de"))

# Convert term to integer
merged["Term_ae"] = merged["Term_ae"].astype(int)
merged["Term_de"] = merged["Term_de"].astype(int)

# Filter where AE term is exactly 1 year after DE term
result = merged[merged["Term_ae"] == merged["Term_de"] + 100]

# Separate students with missing GPA or HS Avg
flagged_df = result[
    (result["AGPA (2 sources)_ae"].isna()) | (result["DE Adm Avg (2 sources)_de"].isna())
]

# Save flagged students to a separate file
flagged_df.to_excel("flagged_students.xlsx", index=False)

# Remove flagged students from analysis
result = result.dropna(subset=["AGPA (2 sources)_ae", "DE Adm Avg (2 sources)_de"])

# Save the cleaned data for analysis
result.to_excel("output_analysis.xlsx", index=False)

# Initialize analysis structure (nested lists)
analysis = [[[[0, 0, 0, 0, 0, 0, 0] for _ in years] for _ in range(2)] for _ in programs]

# Perform analysis
for _, row in result.iterrows():
    prog = row["Program_ae"]  # AE program
    term = int(row["Term_ae"])  # Convert term to int

    # Convert GPA safely (handling multiple values)
    gpa_str = str(row["AGPA (2 sources)_ae"]).strip()
    gpa = float(gpa_str.split()[0]) if gpa_str.replace('.', '', 1).isdigit() else 0.0

    # Convert high school average safely
    hs_avg_str = str(row["DE Adm Avg (2 sources)_de"]).strip()
    hs_avg = float(hs_avg_str.split()[0]) if hs_avg_str.replace('.', '', 1).isdigit() else 0.0

    # Domestic = 0, International = 1
    status = 0 if row["Application Status in Canada_ae"].lower() == "domestic" else 1

    if prog in programs and term in years:
        i = programs.index(prog)  # Find program index
        j = years.index(term)  # Find term index

        # GPA sum and count
        analysis[i][status][j][0] += gpa
        analysis[i][status][j][1] += 1  # Count of students

        # High school average sum
        analysis[i][status][j][2] += hs_avg

        # GPA > 3.5
        if gpa >= 3.5:
            analysis[i][status][j][3] += gpa
            analysis[i][status][j][4] += 1  # Count of students with GPA > 3.5

        # HS Average > 77.78%
        if hs_avg >= 77.78:
            analysis[i][status][j][5] += hs_avg
            analysis[i][status][j][6] += 1  # Count of students with HS avg > 77.78%

# Print analysis results
print("================ Engineering Programs Analysis: ================")
for i, prog in enumerate(programs):
    for j, status in enumerate(["Domestic", "International"]):
        for k, term in enumerate(years):
            data = analysis[i][j][k]
            print(f"{prog} - {status} - {term}:")
            print(f"  GPA Sum: {data[0]}")
            print(f"  # of Students: {data[1]}")
            print(f"  GPA Avg: {data[0] / data[1] if data[1] > 0 else 0}")
            print(f"  High School Avg Sum: {data[2]}")
            print(f"  High School Avg: {data[2] / data[1] if data[1] > 0 else 0}")
            print(f"  GPA >= 3.5 Sum: {data[3]}")
            print(f"  # of Students with GPA >= 3.5: {data[4]}")
            print(f"  Students with GPA >= 3.5 Avg: {data[3] / data[4] if data[4] > 0 else 0}")
            print(f"  HS Avg >= 77.78% Sum: {data[5]}")
            print(f"  # of Students with HS Avg >= 77.78%: {data[6]}")
            print(f"  Students with HS Avg >= 77.78% Avg: {data[5] / data[6] if data[6] > 0 else 0}%")
            print("-" * 40)