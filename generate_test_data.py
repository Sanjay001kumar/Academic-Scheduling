import pandas as pd

data = [
    ["CSE", 1, "Programming", "Dr.Ram", 3, 2],
    ["CSE", 1, "Mathematics", "Dr.Siva", 4, 0],
    ["CSE", 1, "Physics", "Dr.Bala", 3, 2],
    ["CSE", 2, "Data Structures", "Dr.Kumar", 3, 2],
    ["CSE", 2, "DBMS", "Dr.Anita", 3, 2],
    ["IT", 1, "Programming", "Dr.Ram", 3, 2],
    ["IT", 2, "Networks", "Dr.Ravi", 3, 0],
    ["IT", 3, "Machine Learning", "Dr.Ram", 3, 0],
    ["ECE", 2, "Signals", "Dr.Arjun", 3, 0]
]

df = pd.DataFrame(data, columns=["Department", "Year", "Subject", "Faculty", "Theory Hours", "Lab Hours"])

import os
os.makedirs("input", exist_ok=True)
df.to_excel("input/test_input.xlsx", index=False)
print("Created input/test_input.xlsx")
