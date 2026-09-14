"""
Workplace-oriented Excel formula library.

Each item contains:
- formula: formula displayed in the video
- title: short video title
- headers/data: worksheet example
- formula_cell: where the formula is placed
- result: expected displayed result
- explanation: short caption
"""

FORMULAS = [
    {
        "id": "xlookup",
        "title": "Find a value with XLOOKUP",
        "headers": ["Employee", "ID", "Department", "Salary"],
        "rows": [
            ["Ahmed", "E102", "HSE", "8500"],
            ["Sara", "E103", "HR", "9200"],
            ["Ali", "E104", "Finance", "10500"],
        ],
        "formula_cell": "F2",
        "formula": '=XLOOKUP("E103",B2:B4,D2:D4)',
        "result": "9200",
        "explanation": "Find a matching ID and return its salary.",
    },
    {
        "id": "sumifs",
        "title": "Add values with SUMIFS",
        "headers": ["Employee", "Dept", "Overtime Hrs", "Rate", "OT Amount"],
        "rows": [
            ["Ahmed", "HSE", "8", "50", "400"],
            ["Sara", "HR", "5", "50", "250"],
            ["Ali", "HSE", "6", "50", "300"],
        ],
        "formula_cell": "F2",
        "formula": '=SUMIFS(E2:E4,B2:B4,"HSE")',
        "result": "700",
        "explanation": "Add amounts only for HSE employees.",
    },
    {
        "id": "countifs",
        "title": "Count matching records with COUNTIFS",
        "headers": ["Employee", "Dept", "Status"],
        "rows": [
            ["Ahmed", "HSE", "Open"],
            ["Sara", "HR", "Closed"],
            ["Ali", "HSE", "Open"],
            ["Omar", "HSE", "Closed"],
        ],
        "formula_cell": "D2",
        "formula": '=COUNTIFS(B2:B5,"HSE",C2:C5,"Open")',
        "result": "2",
        "explanation": "Count HSE records that are still open.",
    },
    {
        "id": "if",
        "title": "Create a status with IF",
        "headers": ["Employee", "Score", "Status"],
        "rows": [
            ["Ahmed", "92", ""],
            ["Sara", "71", ""],
            ["Ali", "88", ""],
        ],
        "formula_cell": "C2",
        "formula": '=IF(B2>=80,"Pass","Review")',
        "result": "Pass",
        "explanation": "Automatically classify a score.",
    },
    {
        "id": "iferror",
        "title": "Hide errors with IFERROR",
        "headers": ["Sales", "Target", "Achievement"],
        "rows": [
            ["10000", "10000", ""],
            ["7500", "0", ""],
            ["9000", "10000", ""],
        ],
        "formula_cell": "C2",
        "formula": '=IFERROR(A2/B2,0)',
        "result": "100%",
        "explanation": "Return 0 instead of showing a division error.",
    },
    {
        "id": "networkdays",
        "title": "Count working days with NETWORKDAYS",
        "headers": ["Start Date", "End Date", "Working Days"],
        "rows": [
            ["01-Sep-2026", "14-Sep-2026", ""],
            ["07-Sep-2026", "11-Sep-2026", ""],
        ],
        "formula_cell": "C2",
        "formula": "=NETWORKDAYS(A2,B2)",
        "result": "10",
        "explanation": "Count weekdays between two dates.",
    },
    {
        "id": "eomonth",
        "title": "Get month end with EOMONTH",
        "headers": ["Date", "Month End"],
        "rows": [
            ["14-Sep-2026", ""],
            ["02-Oct-2026", ""],
            ["19-Nov-2026", ""],
        ],
        "formula_cell": "B2",
        "formula": "=EOMONTH(A2,0)",
        "result": "30-Sep-2026",
        "explanation": "Return the last day of the same month.",
    },
    {
        "id": "textjoin",
        "title": "Combine cells with TEXTJOIN",
        "headers": ["First Name", "Last Name", "Full Name"],
        "rows": [
            ["Ahmed", "Khan", ""],
            ["Sara", "Ali", ""],
            ["Omar", "Hassan", ""],
        ],
        "formula_cell": "C2",
        "formula": '=TEXTJOIN(" ",TRUE,A2:B2)',
        "result": "Ahmed Khan",
        "explanation": "Combine names without manual typing.",
    },
    {
        "id": "trim",
        "title": "Clean spaces with TRIM",
        "headers": ["Raw Name", "Clean Name"],
        "rows": [
            ["  Ahmed   Khan  ", ""],
            [" Sara Ali ", ""],
            ["  Omar  ", ""],
        ],
        "formula_cell": "B2",
        "formula": "=TRIM(A2)",
        "result": "Ahmed Khan",
        "explanation": "Remove extra spaces from copied data.",
    },
    {
        "id": "round",
        "title": "Round numbers with ROUND",
        "headers": ["Amount", "Rounded"],
        "rows": [
            ["1254.678", ""],
            ["987.456", ""],
            ["42.995", ""],
        ],
        "formula_cell": "B2",
        "formula": "=ROUND(A2,2)",
        "result": "1254.68",
        "explanation": "Round a number to two decimal places.",
    },
    {
        "id": "maxifs",
        "title": "Find a maximum with MAXIFS",
        "headers": ["Employee", "Dept", "Sales"],
        "rows": [
            ["Ahmed", "HSE", "12000"],
            ["Sara", "HR", "15000"],
            ["Ali", "HSE", "18000"],
            ["Omar", "HSE", "14000"],
        ],
        "formula_cell": "D2",
        "formula": '=MAXIFS(C2:C5,B2:B5,"HSE")',
        "result": "18000",
        "explanation": "Find the highest HSE sales value.",
    },
    {
        "id": "unique",
        "title": "Remove duplicates with UNIQUE",
        "headers": ["Department", "Unique Departments"],
        "rows": [
            ["HSE", ""],
            ["HR", ""],
            ["HSE", ""],
            ["Finance", ""],
        ],
        "formula_cell": "B2",
        "formula": "=UNIQUE(A2:A5)",
        "result": "HSE • HR • Finance",
        "explanation": "Create a unique list automatically.",
    },
    {
        "id": "filter",
        "title": "Filter records with FILTER",
        "headers": ["Employee", "Status", "Department"],
        "rows": [
            ["Ahmed", "Open", "HSE"],
            ["Sara", "Closed", "HR"],
            ["Ali", "Open", "Finance"],
        ],
        "formula_cell": "D2",
        "formula": '=FILTER(A2:C4,B2:B4="Open")',
        "result": "Ahmed • Ali",
        "explanation": "Show only rows matching a condition.",
    },
    {
        "id": "left",
        "title": "Extract text with LEFT",
        "headers": ["Employee ID", "Prefix"],
        "rows": [
            ["HSE-1024", ""],
            ["HR-2045", ""],
            ["FIN-3100", ""],
        ],
        "formula_cell": "B2",
        "formula": "=LEFT(A2,3)",
        "result": "HSE",
        "explanation": "Extract the first characters from an ID.",
    },
    {
        "id": "right",
        "title": "Extract numbers with RIGHT",
        "headers": ["Reference", "Number"],
        "rows": [
            ["INV-2026-1045", ""],
            ["INV-2026-1088", ""],
            ["INV-2026-1102", ""],
        ],
        "formula_cell": "B2",
        "formula": "=RIGHT(A2,4)",
        "result": "1045",
        "explanation": "Extract the last four characters.",
    },
    {
        "id": "workday",
        "title": "Add working days with WORKDAY",
        "headers": ["Start Date", "Days", "Due Date"],
        "rows": [
            ["14-Sep-2026", "5", ""],
            ["15-Sep-2026", "10", ""],
        ],
        "formula_cell": "C2",
        "formula": "=WORKDAY(A2,B2)",
        "result": "21-Sep-2026",
        "explanation": "Calculate a due date excluding weekends.",
    },
    {
        "id": "averageif",
        "title": "Average matching values with AVERAGEIF",
        "headers": ["Dept", "Score", "Average HSE"],
        "rows": [
            ["HSE", "80", ""],
            ["HR", "70", ""],
            ["HSE", "90", ""],
            ["HSE", "100", ""],
        ],
        "formula_cell": "C2",
        "formula": '=AVERAGEIF(A2:A5,"HSE",B2:B5)',
        "result": "90",
        "explanation": "Average only the HSE scores.",
    },
]

def get_daily_formulas(date_obj, count=3):
    """Deterministically rotate through the library by calendar day."""
    start = (date_obj - __import__("datetime").date(2026, 1, 1)).days
    n = len(FORMULAS)
    return [FORMULAS[(start * count + i) % n] for i in range(count)]
