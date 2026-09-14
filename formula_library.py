from datetime import date

VIDEOS = [
    {
        "id": "checkbox_status",
        "title": "Create checkboxes that update status",
        "hook": "ADD CHECKBOXES IN EXCEL — ONE BY ONE",
        "headers": ["Bills", "Checkbox", "Status"],
        "rows": [
            ["Electricity", False, "Paid"],
            ["Gas", False, "Paid"],
            ["Water", False, "Paid"],
            ["Insurance", False, "Unpaid"],
            ["Internet", False, "Unpaid"],
            ["Mobile Phones", False, "Unpaid"],
            ["Debit Card", False, "Paid"],
            ["Washing Machines", False, "Paid"],
            ["Heaters", False, "Unpaid"],
        ],
        "voice": "Want to add clickable checkboxes to an Excel list? First, insert a checkbox in the first row. Then copy it down the column. Now each checkbox can control the status. Check a box and watch the status update automatically. Follow Learn Verse for more Excel tips.",
        "caption_steps": [
            (0.0, 4.0, "ADD CHECKBOXES\nONE BY ONE"),
            (4.0, 8.0, "IN EXCEL JUST\nADD A CHECKBOX"),
            (8.0, 12.0, "COPY IT DOWN\nYOUR COLUMN"),
            (12.0, 17.0, "STATUS\nUPDATES AUTOMATICALLY"),
            (17.0, 22.0, "CHECK THE BOX\nAND WATCH"),
            (22.0, 27.0, "YOUR SHEET\nUPDATES"),
            (27.0, 30.0, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],
        "type": "checkbox"
    },
    {
        "id": "xlookup",
        "title": "Find a salary with XLOOKUP",
        "hook": "FIND ANY SALARY FROM AN ID — FAST",
        "headers": ["Employee", "ID", "Department", "Salary"],
        "rows": [
            ["Ahmed", "E102", "HSE", 8500],
            ["Sara", "E103", "HR", 9200],
            ["Ali", "E104", "Finance", 10500],
            ["Omar", "E105", "IT", 11000],
        ],
        "formula": '=XLOOKUP("E103",B2:B5,D2:D5)',
        "result": "9200",
        "voice": "Need a salary from an employee ID? Use XLOOKUP. Select the ID you need, then enter the XLOOKUP formula. Excel searches the ID column and returns the matching salary. Here the result is nine thousand two hundred. Follow Learn Verse for more Excel tips.",
        "caption_steps": [
            (0.0, 3.5, "FIND A SALARY\nFROM AN ID"),
            (3.5, 8.0, "SELECT THE\nEMPLOYEE ID"),
            (8.0, 15.0, "TYPE XLOOKUP\nTO FIND IT"),
            (15.0, 22.0, "EXCEL RETURNS\nTHE MATCH"),
            (22.0, 27.0, "RESULT = 9200"),
            (27.0, 30.0, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],
        "type": "formula"
    },
]

def get_daily_videos(run_date: date, count=3):
    # Rotation is deterministic; repeats only after the library is exhausted.
    offset = ((run_date - date(2026, 1, 1)).days * count) % len(VIDEOS)
    return [VIDEOS[(offset+i) % len(VIDEOS)] for i in range(count)]
