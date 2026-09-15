from __future__ import annotations

from datetime import date


# ============================================================
# LEARN VERSE - EXCEL SHORTS CONTENT LIBRARY
# ============================================================
#
# The generator automatically selects 3 different topics
# according to the date.
#
# Each topic contains:
#   - Excel data
#   - real Excel formula
#   - voice script
#   - synchronized captions
#   - YouTube title
#   - YouTube description
#   - hashtags
#
# ============================================================


VIDEOS = [

    # ========================================================
    # 01 - XLOOKUP
    # ========================================================
    {
        "id": "xlookup",
        "title": "XLOOKUP",
        "type": "formula",

        "headers": [
            "Employee",
            "Employee ID",
            "Department",
            "Salary"
        ],

        "rows": [
            ["Ahmed", "E102", "HSE", 8500],
            ["Sara", "E103", "HR", 9200],
            ["Ali", "E104", "Finance", 10500],
            ["Omar", "E105", "IT", 11000],
            ["Hina", "E106", "Admin", 7800],
        ],

        "input": "E103",

        "formula": '=XLOOKUP(E103,B5:B9,D5:D9,"Not found")',

        "result": "9200",

        "voice": (
            "Need a salary from an employee ID? "
            "Use XLOOKUP. "
            "Type the ID, enter the formula, and Excel returns "
            "the matching salary instantly. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "FIND A VALUE\nFROM AN ID"),
            (4, 9, "TYPE THE\nEMPLOYEE ID"),
            (9, 17, "USE XLOOKUP\nTO FIND IT"),
            (17, 24, "EXCEL RETURNS\nTHE MATCH"),
            (24, 27, "RESULT = 9200"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "Master XLOOKUP in Excel in 30 Seconds!"
        ),

        "youtube_description": (
            "Learn how to use XLOOKUP in Microsoft Excel "
            "with a practical employee salary example.\n\n"
            "Perfect for beginners learning Excel formulas "
            "and lookup functions.\n\n"
            "Follow Learn Verse for more Excel tips, formulas "
            "and shortcuts."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#XLOOKUP",
            "#MicrosoftExcel",
            "#ExcelTutorial",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 02 - IF
    # ========================================================
    {
        "id": "if_status",
        "title": "IF Function",
        "type": "formula",

        "headers": [
            "Task",
            "Score",
            "Target",
            "Status"
        ],

        "rows": [
            ["Safety", 85, 70, ""],
            ["Quality", 62, 70, ""],
            ["Training", 91, 70, ""],
            ["Inspection", 74, 70, ""],
            ["Audit", 55, 70, ""],
        ],

        "input": "85",

        "formula": '=IF(B5>=C5,"PASS","FAIL")',

        "result": "PASS",

        "voice": (
            "Turn a score into a clear status with IF. "
            "Compare the score with your target, "
            "then return PASS or FAIL automatically. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "TURN SCORES INTO\nPASS OR FAIL"),
            (4, 9, "START WITH\nYOUR SCORE"),
            (9, 17, "USE THE IF\nFUNCTION"),
            (17, 24, "EXCEL DECIDES\nAUTOMATICALLY"),
            (24, 27, "RESULT = PASS"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "Excel IF Function Explained in 30 Seconds!"
        ),

        "youtube_description": (
            "Learn how to use the IF function in Microsoft Excel "
            "to automatically return PASS or FAIL based on a score.\n\n"
            "A simple and useful Excel formula for reports, "
            "assessments and business data."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#IFFunction",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 03 - SUMIF
    # ========================================================
    {
        "id": "sumif",
        "title": "SUMIF",
        "type": "formula",

        "headers": [
            "Product",
            "Region",
            "Sales",
            "Salesperson"
        ],

        "rows": [
            ["Phone", "UAE", 1200, "Ahmed"],
            ["Laptop", "UAE", 2500, "Sara"],
            ["Tablet", "KSA", 900, "Ali"],
            ["Monitor", "UAE", 1500, "Omar"],
            ["Keyboard", "KSA", 400, "Hina"],
        ],

        "input": "UAE",

        "formula": '=SUMIF(B5:B9,"UAE",C5:C9)',

        "result": "5200",

        "voice": (
            "Need total sales for one region? "
            "SUMIF can do it in one formula. "
            "Set the region as your condition, "
            "and Excel adds only matching sales. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "TOTAL SALES\nBY REGION"),
            (4, 9, "CHOOSE YOUR\nREGION"),
            (9, 17, "USE SUMIF\nAS THE FORMULA"),
            (17, 24, "EXCEL ADDS\nMATCHING SALES"),
            (24, 27, "RESULT = 5200"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "SUMIF in Excel Made Easy!"
        ),

        "youtube_description": (
            "Learn how to use SUMIF in Microsoft Excel "
            "to calculate sales for a specific region.\n\n"
            "This practical example shows how Excel can "
            "automatically add only matching values."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#SUMIF",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 04 - COUNTIF
    # ========================================================
    {
        "id": "countif",
        "title": "COUNTIF",
        "type": "formula",

        "headers": [
            "Employee",
            "Department",
            "Status",
            "Priority"
        ],

        "rows": [
            ["Ahmed", "HSE", "Done", "High"],
            ["Sara", "HR", "Pending", "Medium"],
            ["Ali", "Finance", "Done", "High"],
            ["Omar", "IT", "Done", "Low"],
            ["Hina", "Admin", "Pending", "Medium"],
            ["Zaid", "Operations", "Done", "High"],
        ],

        "input": "Done",

        "formula": '=COUNTIF(C5:C10,"Done")',

        "result": "4",

        "voice": (
            "Want to count completed tasks instantly? "
            "COUNTIF counts only the cells that match your condition. "
            "Use it for tasks, attendance, sales, or status lists. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "COUNT COMPLETED\nTASKS FAST"),
            (4, 9, "SET YOUR\nSTATUS"),
            (9, 17, "USE COUNTIF\nTO COUNT IT"),
            (17, 24, "EXCEL COUNTS\nMATCHES"),
            (24, 27, "RESULT = 4"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "COUNTIF in Excel: Count Tasks Instantly!"
        ),

        "youtube_description": (
            "Learn how to use COUNTIF in Microsoft Excel "
            "to count completed tasks automatically.\n\n"
            "COUNTIF is useful for task tracking, attendance, "
            "reports and business dashboards."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#COUNTIF",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 05 - IFERROR
    # ========================================================
    {
        "id": "iferror",
        "title": "IFERROR",
        "type": "formula",

        "headers": [
            "Item",
            "Total Cost",
            "Quantity",
            "Unit Cost"
        ],

        "rows": [
            ["Keyboard", 160, 2, ""],
            ["Mouse", 135, 3, ""],
            ["Monitor", 700, 1, ""],
            ["Headset", 240, 4, ""],
            ["Laptop", 2500, 1, ""],
        ],

        "input": "Keyboard",

        "formula": '=IFERROR(B5/C5,"Check values")',

        "result": "80",

        "voice": (
            "Stop ugly Excel errors from showing in your report. "
            "Wrap your formula with IFERROR and display a clean "
            "message when something goes wrong. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "HIDE EXCEL\nERRORS"),
            (4, 9, "START WITH\nYOUR FORMULA"),
            (9, 17, "WRAP IT WITH\nIFERROR"),
            (17, 24, "SHOW A CLEAN\nRESULT"),
            (24, 27, "RESULT = 80"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "IFERROR in Excel: Hide Errors the Smart Way!"
        ),

        "youtube_description": (
            "Learn how to use IFERROR in Microsoft Excel "
            "to replace formula errors with a clean result "
            "or helpful message."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#IFERROR",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 06 - AVERAGE
    # ========================================================
    {
        "id": "average",
        "title": "AVERAGE",
        "type": "formula",

        "headers": [
            "Employee",
            "Score",
            "Department",
            "Review"
        ],

        "rows": [
            ["Ahmed", 80, "HSE", "Good"],
            ["Sara", 90, "HR", "Excellent"],
            ["Ali", 70, "Finance", "Good"],
            ["Omar", 100, "IT", "Excellent"],
            ["Hina", 85, "Admin", "Good"],
        ],

        "input": "Scores",

        "formula": '=AVERAGE(B5:B9)',

        "result": "85",

        "voice": (
            "Need the average of a score list? "
            "Use AVERAGE and select the cells. "
            "Excel calculates the mean instantly. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "GET THE\nAVERAGE FAST"),
            (4, 9, "SELECT YOUR\nSCORES"),
            (9, 17, "USE AVERAGE\nAS THE FORMULA"),
            (17, 24, "EXCEL CALCULATES\nTHE MEAN"),
            (24, 27, "RESULT = 85"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "AVERAGE Formula in Excel in 30 Seconds!"
        ),

        "youtube_description": (
            "Learn how to calculate the average of a list "
            "using the AVERAGE function in Microsoft Excel."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#AVERAGE",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 07 - MAX
    # ========================================================
    {
        "id": "max",
        "title": "MAX",
        "type": "formula",

        "headers": [
            "Employee",
            "Sales",
            "Region",
            "Target"
        ],

        "rows": [
            ["Ahmed", 1200, "UAE", 1500],
            ["Sara", 2100, "UAE", 1800],
            ["Ali", 1800, "KSA", 1800],
            ["Omar", 1600, "UAE", 1700],
            ["Hina", 1950, "KSA", 1800],
        ],

        "input": "Sales",

        "formula": '=MAX(B5:B9)',

        "result": "2100",

        "voice": (
            "Need the highest number in a list? "
            "MAX returns the largest value instantly. "
            "Select your range and Excel finds the top result. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "FIND THE\nHIGHEST VALUE"),
            (4, 9, "SELECT YOUR\nNUMBER RANGE"),
            (9, 17, "USE MAX\nAS THE FORMULA"),
            (17, 24, "EXCEL FINDS\nTHE TOP VALUE"),
            (24, 27, "RESULT = 2100"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "Find the Highest Value with MAX in Excel!"
        ),

        "youtube_description": (
            "Learn how to use the MAX function in Microsoft Excel "
            "to instantly find the highest number in a range."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#MAX",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 08 - MIN
    # ========================================================
    {
        "id": "min",
        "title": "MIN",
        "type": "formula",

        "headers": [
            "Product",
            "Price",
            "Category",
            "Stock"
        ],

        "rows": [
            ["Phone", 1800, "Mobile", 15],
            ["Tablet", 900, "Mobile", 22],
            ["Watch", 650, "Accessories", 35],
            ["Laptop", 2500, "Computer", 8],
            ["Keyboard", 120, "Accessories", 50],
        ],

        "input": "Price",

        "formula": '=MIN(B5:B9)',

        "result": "120",

        "voice": (
            "Need the lowest price in a list? "
            "MIN finds the smallest number instantly. "
            "Select your range and let Excel do the work. "
            "Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "FIND THE\nLOWEST PRICE"),
            (4, 9, "SELECT YOUR\nPRICE RANGE"),
            (9, 17, "USE MIN\nAS THE FORMULA"),
            (17, 24, "EXCEL FINDS\nTHE LOWEST"),
            (24, 27, "RESULT = 120"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "Find the Lowest Price with MIN in Excel!"
        ),

        "youtube_description": (
            "Learn how to use the MIN function in Microsoft Excel "
            "to find the smallest number in a data range."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#MIN",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 09 - SUM
    # ========================================================
    {
        "id": "sum",
        "title": "SUM",
        "type": "formula",

        "headers": [
            "Expense",
            "Amount",
            "Category",
            "Month"
        ],

        "rows": [
            ["Rent", 3000, "Office", "September"],
            ["Food", 1200, "Operations", "September"],
            ["Fuel", 500, "Transport", "September"],
            ["Bills", 800, "Utilities", "September"],
            ["Supplies", 650, "Office", "September"],
        ],

        "input": "Expenses",

        "formula": '=SUM(B5:B9)',

        "result": "6150",

        "voice": (
            "Add an entire column in seconds with SUM. "
            "Select the numbers, enter SUM, and Excel gives you "
            "the total. Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "ADD A COLUMN\nIN SECONDS"),
            (4, 9, "SELECT YOUR\nNUMBERS"),
            (9, 17, "USE SUM\nAS THE FORMULA"),
            (17, 24, "EXCEL CALCULATES\nTHE TOTAL"),
            (24, 27, "RESULT = 6150"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "SUM Formula in Excel: Add Numbers Fast!"
        ),

        "youtube_description": (
            "Learn how to use the SUM function in Microsoft Excel "
            "to quickly calculate totals from a list of numbers."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#SUM",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 10 - TEXTJOIN
    # ========================================================
    {
        "id": "textjoin",
        "title": "TEXTJOIN",
        "type": "formula",

        "headers": [
            "First Name",
            "Last Name",
            "Department",
            "Email"
        ],

        "rows": [
            ["Ahmed", "Khan", "HSE", ""],
            ["Sara", "Ali", "HR", ""],
            ["Omar", "Hassan", "IT", ""],
            ["Hina", "Malik", "Admin", ""],
        ],

        "input": "Ahmed + Khan",

        "formula": '=TEXTJOIN(" ",TRUE,A5:B5)',

        "result": "Ahmed Khan",

        "voice": (
            "Combine text from multiple cells without manual typing. "
            "TEXTJOIN can join names, labels, or descriptions "
            "with a separator. Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "COMBINE TEXT\nAUTOMATICALLY"),
            (4, 9, "SELECT THE\nTEXT CELLS"),
            (9, 17, "USE TEXTJOIN\nTO COMBINE"),
            (17, 24, "EXCEL JOINS\nTHE VALUES"),
            (24, 27, "RESULT = AHMED KHAN"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "TEXTJOIN in Excel: Combine Text Automatically!"
        ),

        "youtube_description": (
            "Learn how to use TEXTJOIN in Microsoft Excel "
            "to combine text from multiple cells automatically."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#TEXTJOIN",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 11 - INDEX + MATCH
    # ========================================================
    {
        "id": "index_match",
        "title": "INDEX + MATCH",
        "type": "formula",

        "headers": [
            "Employee",
            "Employee ID",
            "Salary",
            "Department"
        ],

        "rows": [
            ["Ahmed", "E102", 8500, "HSE"],
            ["Sara", "E103", 9200, "HR"],
            ["Ali", "E104", 10500, "Finance"],
            ["Omar", "E105", 11000, "IT"],
            ["Hina", "E106", 7800, "Admin"],
        ],

        "input": "E103",

        "formula": '=INDEX(C5:C9,MATCH(E103,B5:B9,0))',

        "result": "9200",

        "voice": (
            "Here is a powerful lookup combination. "
            "MATCH finds the position, and INDEX returns the value "
            "from that position. Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "POWERFUL\nLOOKUP COMBO"),
            (4, 9, "TYPE THE\nEMPLOYEE ID"),
            (9, 17, "MATCH FINDS\nTHE POSITION"),
            (17, 24, "INDEX RETURNS\nTHE VALUE"),
            (24, 27, "RESULT = 9200"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "INDEX + MATCH in Excel Explained Fast!"
        ),

        "youtube_description": (
            "Learn how INDEX and MATCH work together in Microsoft Excel "
            "to find and return a matching value from a table."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#INDEXMATCH",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },


    # ========================================================
    # 12 - REAL EXCEL CHECKBOX
    # ========================================================
    {
        "id": "checkbox_status",
        "title": "Excel Checkbox",
        "type": "checkbox",

        "headers": [
            "Task",
            "Responsible",
            "Done",
            "Status"
        ],

        "rows": [
            ["Safety Inspection", "Ahmed", None, ""],
            ["Daily Report", "Sara", None, ""],
            ["Toolbox Talk", "Ali", None, ""],
            ["Permit Check", "Omar", None, ""],
            ["Training", "Hina", None, ""],
            ["Audit", "Zaid", None, ""],
        ],

        "voice": (
            "Want clickable checkboxes in Excel? "
            "Insert a checkbox, link it to a cell, "
            "and let an IF formula turn TRUE or FALSE "
            "into a clean status. Follow Learn Verse for more Excel tips."
        ),

        "captions": [
            (0, 4, "ADD CLICKABLE\nCHECKBOXES"),
            (4, 9, "CLICK A\nCHECKBOX"),
            (9, 17, "LINK IT TO\nA CELL"),
            (17, 24, "IF FORMULA\nUPDATES STATUS"),
            (24, 27, "DONE = TRUE"),
            (27, 30, "FOLLOW FOR MORE\nEXCEL TIPS"),
        ],

        "youtube_title": (
            "Add Real Clickable Checkboxes in Excel!"
        ),

        "youtube_description": (
            "Learn how to use real Excel checkboxes with linked cells "
            "and an IF formula to automatically display task status."
        ),

        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#Checkbox",
            "#MicrosoftExcel",
            "#ExcelTutorial",
            "#LearnExcel",
            "#LearnVerse",
        ],
    },
]


# ============================================================
# DAILY ROTATION
# ============================================================

def get_daily_videos(run_date: date, count: int = 3):
    """
    Return a deterministic set of videos for a specific date.

    The same date always produces the same 3 topics.
    Different days rotate through the complete library.
    """

    if count < 1:
        return []

    if count > len(VIDEOS):
        count = len(VIDEOS)

    # Rotate through the library using the calendar date.
    base_date = date(2026, 1, 1)
    days_since_start = (run_date - base_date).days

    start_index = (days_since_start * 3) % len(VIDEOS)

    selected = []

    for i in range(count):
        index = (start_index + i) % len(VIDEOS)
        selected.append(VIDEOS[index])

    return selected


# ============================================================
# YOUTUBE METADATA HELPERS
# ============================================================

def get_youtube_title(video):
    """
    Return the YouTube title for a video.
    """

    return video.get(
        "youtube_title",
        f"{video.get('title', 'Excel Tip')} | Learn Verse"
    )


def get_youtube_description(video):
    """
    Return the YouTube description.
    """

    description = video.get(
        "youtube_description",
        "Learn Microsoft Excel with Learn Verse."
    )

    hashtags = video.get("hashtags", [])

    if hashtags:
        description = (
            description.rstrip()
            + "\n\n"
            + " ".join(hashtags)
        )

    return description


def get_youtube_hashtags(video):
    """
    Return hashtags as a list.
    """

    return list(video.get("hashtags", []))


def get_metadata(video):
    """
    Return all YouTube metadata in one dictionary.
    """

    return {
        "title": get_youtube_title(video),
        "description": get_youtube_description(video),
        "hashtags": get_youtube_hashtags(video),
        "topic": video.get("title", ""),
        "id": video.get("id", ""),
    }