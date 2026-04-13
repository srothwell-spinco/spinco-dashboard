import os
import pandas as pd

DATA_DIR = "data/incoming"

# Find CSV or Excel files in the folder
files = [
    f for f in os.listdir(DATA_DIR)
    if f.lower().endswith((".csv", ".xlsx", ".xls"))
]

if not files:
    raise SystemExit("No data files found in data/incoming")

# Use the first file (for now)
file_path = os.path.join(DATA_DIR, files[0])
print("Reading file:", file_path)

# Load the file
if file_path.lower().endswith(".csv"):
    df = pd.read_csv(file_path)
else:
    df = pd.read_excel(file_path)

print("\nCOLUMNS:")
for col in df.columns:
    print(" -", col)

print("\nFIRST 5 ROWS:")
print(df.head())

import os
import pandas as pd

DATA_DIR = "data/incoming"

# Find CSV or Excel files in the folder
files = [
    f for f in os.listdir(DATA_DIR)
    if f.lower().endswith((".csv", ".xlsx", ".xls"))
]

if not files:
    raise SystemExit("No data files found in data/incoming")

# Use the first file (for now)
file_path = os.path.join(DATA_DIR, files[0])
print("Reading file:", file_path)

# Load the file
if file_path.lower().endswith(".csv"):
    df = pd.read_csv(file_path)
else:
    df = pd.read_excel(file_path)

print("\nCOLUMNS:")
for col in df.columns:
    print(" -", col)

print("\nFIRST 5 ROWS:")
print(df.head())

