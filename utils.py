import pandas as pd

def export_to_excel(timetable, path):
    """
    Exports a multi-department timetable to an Excel file with multiple sheets.
    timetable[dept][year][day][period]
    """
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for dept, dept_data in timetable.items():
            for year, year_data in dept_data.items():
                
                periods = set()
                for day_map in year_data.values():
                    periods.update(day_map.keys())

                def period_key(p):
                    try:
                        return int(p.lstrip('P'))
                    except Exception:
                        return p

                periods = sorted(list(periods), key=period_key)
                columns = ["Day"] + periods
                rows = []

                for day, day_map in year_data.items():
                    row = [day]
                    for p in periods:
                        val = day_map.get(p)
                        row.append(val if val else "Free")
                    rows.append(row)

                df = pd.DataFrame(rows, columns=columns)
                
                # Sheet name e.g., "CSE - Year 1"
                sheet_name = f"{dept} - Y{year}"
                # Ensure sheet name length is ok
                sheet_name = sheet_name[:31]
                
                df.to_excel(writer, index=False, sheet_name=sheet_name)
    
