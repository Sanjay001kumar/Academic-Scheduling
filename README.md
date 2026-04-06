# Intelligent Academic Scheduling Engine

An AI-Based Smart College Timetable Generator written in Python, using Constraint Programming (Google OR-Tools) to automatically schedule a conflict-free weekly timetable for multiple departments.

## Features
- **Constraint-Based Optimization**: Leverages Google OR-Tools to solve complex scheduling logic.
- **Multidimensional Constraints**: Prevents faculty overlap, groups continuous lab sessions, prioritizes evenly distributed workloads without gaps, and ensures appropriate credit hour fulfilling per subject.
- **Web Interface**: Easy-to-use Flask web app for uploading inputs and displaying a highly visual output timetable.
- **Export to Excel**: Generates downloadable excel sheets grouped gracefully by department.

## Project Structure
- `app.py`: Flask backend serving the UI and handling file uploads/downloads.
- `scheduler.py`: The core algorithmic constraint solver mapping out the logic with OR-Tools.
- `utils.py`: Helper functions, specifically exporting multi-level dictionary timetables to Excel sheets.
- `generate_test_data.py`: A quick helper script to quickly generate a sample Input Excel sheet.
- `input/test_input.xlsx`: Example data containing Department, Year, Subject, Faculty, Theory Hours, Lab Hours.
- `templates/`: Contains HTML structure for the web interface styling (`index.html` & `result.html`).

## Installation

1. Make sure you have python installed (Python 3.8+ recommended).
2. Install the necessary dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Create your input Excel sheet or simply run `python generate_test_data.py` to auto-create a sample file inside `input/`.
2. Start the web server:
   ```bash
   python app.py
   ```
3. Open a browser and navigate to: http://127.0.0.1:5000
4. Upload your Excel Input and hit "Generate Timetable".
5. Browse your results right on your screen, and click "Download Full Excel" for a highly readable multi-sheet timetable.

## Constraints Ensured
- A faculty member never teaches two simultaneous classes.
- Each department and year gets its own specific conflict-free timetable.
- Lab hours are automatically assigned continuous periods.
- Preferred tight schedules minimizing empty slots across the schedule matrix.
