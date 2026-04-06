import pandas as pd

def fix_input_data(filename):
    # First, let's remove any existing padding from previous runs to avoid duplicating them
    # Filtering out existing padded rows
    df1 = pd.read_excel(filename)
    df1 = df1[df1["Subject"] != "Library/Seminar"] # Filter out old padding
    
    new_rows = []
    
    # Calculate current hours per department and year
    df1["Theory Hours"] = pd.to_numeric(df1["Theory Hours"], errors="coerce").fillna(0)
    df1["Lab Hours"] = pd.to_numeric(df1["Lab Hours"], errors="coerce").fillna(0)
    
    groups = df1.groupby(["Department", "Year"])
    for (dept, year), group in groups:
        total = group["Theory Hours"].sum() + group["Lab Hours"].sum()
        if total < 30:
            missing = 30 - total
            new_rows.append({
                "Department": dept,
                "Year": year,
                "Subject": "Library/Seminar",
                "Faculty": f"Staff_{dept}_{year}",
                "Theory Hours": int(missing),
                "Lab Hours": 0
            })
            
    if new_rows:
        new_df = pd.DataFrame(new_rows)
        df1 = pd.concat([df1, new_df], ignore_index=True)
        
    df1.to_excel(filename, index=False)
    print(f"Fixed {filename}, added {len(new_rows)} rows to fill missing hours.")

if __name__ == "__main__":
    fix_input_data("college_timetable_input.xlsx")
    try:
        fix_input_data("input/test_input.xlsx")
    except Exception:
        pass
