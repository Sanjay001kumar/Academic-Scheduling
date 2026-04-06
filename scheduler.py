import pandas as pd
import collections
import math
from ortools.sat.python import cp_model

DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday"]
PERIODS = ["P1","P2","P3","P4","P5","P6"]

def generate_timetable(df):

    model = cp_model.CpModel()

    num_days = len(DAYS)
    num_periods = len(PERIODS)

    courses = []

    # Clean numeric columns
    df["Theory Hours"] = pd.to_numeric(df["Theory Hours"], errors="coerce")
    df["Lab Hours"] = pd.to_numeric(df["Lab Hours"], errors="coerce")
    df = df.dropna(subset=["Theory Hours","Lab Hours"])

    for idx,row in df.iterrows():

        courses.append({
            "idx": idx,
            "dept": str(row["Department"]),
            "year": str(row["Year"]),
            "subject": str(row["Subject"]),
            "faculty": str(row["Faculty"]),
            "theory": int(row["Theory Hours"]),
            "lab": int(row["Lab Hours"])
        })

    # VARIABLES
    x_theory = {}
    x_lab = {}

    for c in courses:
        c_idx = c["idx"]

        for d in range(num_days):
            for p in range(num_periods):

                x_theory[(c_idx,d,p)] = model.NewBoolVar(f"t_{c_idx}_{d}_{p}")

                if p < num_periods-1:
                    x_lab[(c_idx,d,p)] = model.NewBoolVar(f"l_{c_idx}_{d}_{p}")

    # Occupancy function
    def get_occupancy(c_idx,d,p):

        occ = [x_theory[(c_idx,d,p)]]

        if p < num_periods-1:
            occ.append(x_lab[(c_idx,d,p)])

        if p > 0:
            occ.append(x_lab[(c_idx,d,p-1)])

        return sum(occ)

    groups = collections.defaultdict(list)
    faculty_dict = collections.defaultdict(list)

    for c in courses:

        c_idx = c["idx"]
        dept = c["dept"]
        year = c["year"]
        faculty = c["faculty"]

        groups[(dept,year)].append(c_idx)
        faculty_dict[faculty].append(c_idx)

        # THEORY HOURS
        model.Add(
            sum(x_theory[(c_idx,d,p)]
            for d in range(num_days)
            for p in range(num_periods))
            == c["theory"]
        )

        # LAB HOURS
        lab_blocks = c["lab"] // 2

        model.Add(
            sum(x_lab[(c_idx,d,p)]
            for d in range(num_days)
            for p in range(num_periods-1))
            == lab_blocks
        )

        # Spread subjects
        max_t_per_day = max(1,math.ceil(c["theory"]/num_days))

        for d in range(num_days):

            model.Add(
                sum(x_theory[(c_idx,d,p)]
                for p in range(num_periods))
                <= max_t_per_day
            )

    # ONE SUBJECT PER SLOT PER CLASS
    for g,g_courses in groups.items():

        for d in range(num_days):
            for p in range(num_periods):

                model.Add(
                    sum(get_occupancy(c_idx,d,p)
                    for c_idx in g_courses)
                    <= 1
                )

    # FACULTY CONFLICT
    for f,f_courses in faculty_dict.items():

        for d in range(num_days):
            for p in range(num_periods):

                model.Add(
                    sum(get_occupancy(c_idx,d,p)
                    for c_idx in f_courses)
                    <= 1
                )

    # Removed objective forcing P6 to be free. The model will naturally distribute classes, 
    # preventing quadratic penalties from squishing schedules early in the day.

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 15

    status = solver.Solve(model)

    timetable = {}

    for g in groups.keys():

        dept,year = g

        if dept not in timetable:
            timetable[dept] = {}

        timetable[dept][year] = {}

        for d_name in DAYS:

            timetable[dept][year][d_name] = {
                p:"Free" for p in PERIODS
            }

    if status in [cp_model.OPTIMAL,cp_model.FEASIBLE]:

        for c in courses:

            c_idx = c["idx"]
            dept = c["dept"]
            year = c["year"]
            subj = c["subject"]
            fac = c["faculty"]

            for d in range(num_days):

                day = DAYS[d]

                for p in range(num_periods):

                    if solver.Value(x_theory[(c_idx,d,p)]) == 1:

                        timetable[dept][year][day][PERIODS[p]] = f"{subj} ({fac})"

                    if p < num_periods-1:

                        if solver.Value(x_lab[(c_idx,d,p)]) == 1:

                            lab_text = f"{subj} ({fac}) - LAB"

                            timetable[dept][year][day][PERIODS[p]] = lab_text
                            timetable[dept][year][day][PERIODS[p+1]] = lab_text

    return timetable


def print_timetable(timetable):

    for dept in timetable:

        for year in timetable[dept]:

            print(f"\n{dept} - Year {year}")

            header = ["Day"] + PERIODS
            print("\t".join(header))

            for day in DAYS:

                row = [day]

                for p in PERIODS:

                    row.append(
                        timetable[dept][year][day][p]
                    )

                print("\t".join(row))


# MAIN
if __name__ == "__main__":

    df = pd.read_excel("college_timetable_input.xlsx")

    timetable = generate_timetable(df)

    if timetable:
        print_timetable(timetable)
        from utils import export_to_excel
        export_to_excel(timetable, "outputs/final_timetable.xlsx")
        print("\nTimetable successfully generated and saved to outputs/final_timetable.xlsx")
    else:
        print("Failed to generate timetable. It may be infeasible.")