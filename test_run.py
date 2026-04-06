import pandas as pd
from scheduler import generate_timetable

df = pd.read_excel("input/test_input.xlsx")
timetable = generate_timetable(df)

if not timetable:
    print("Timetable generation failed.")
else:
    for dept, y_map in timetable.items():
        for year, d_map in y_map.items():
            print(f"--- {dept} Year {year} ---")
            for day, calls in d_map.items():
                print(f"{day}: {calls}")
    print("Success")
