from datetime import date

FORMULAS = [
    {
        "id": "if",
        "problem": "Need Excel to mark employees as PASS or REVIEW?",
        "title": "Automatically mark PASS or REVIEW",
        "headers": ["Employee", "Score", "Status"],
        "rows": [["Ahmed", 92, ""], ["Sara", 71, ""], ["Ali", 88, ""]],
        "formula": '=IF(B2>=80,"Pass","Review")',
        "result": "Pass",
        "explanation": "IF checks the score and returns the correct status.",
        "voice": "Need Excel to mark employee scores automatically? Use the IF formula. If the score is eighty or higher, Excel returns Pass. Otherwise, it returns Review. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "xlookup",
        "problem": "Need to find an employee salary from an ID?",
        "title": "Find data instantly with XLOOKUP",
        "headers": ["Employee", "ID", "Department", "Salary"],
        "rows": [["Ahmed", "E102", "HSE", 8500], ["Sara", "E103", "HR", 9200], ["Ali", "E104", "Finance", 10500]],
        "formula": '=XLOOKUP("E103",B2:B4,D2:D4)',
        "result": "9200",
        "explanation": "XLOOKUP finds the ID and returns the matching salary.",
        "voice": "Need to find an employee salary from an ID? Use XLOOKUP. Excel searches the ID column, finds E103, and returns the matching salary. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "sumifs",
        "problem": "Need the total overtime amount for one department?",
        "title": "Total overtime with SUMIFS",
        "headers": ["Employee", "Dept", "OT Hours", "Rate", "OT Amount"],
        "rows": [["Ahmed", "HSE", 8, 50, 400], ["Sara", "HR", 5, 50, 250], ["Ali", "HSE", 6, 50, 300]],
        "formula": '=SUMIFS(E2:E4,B2:B4,"HSE")',
        "result": "700",
        "explanation": "SUMIFS adds only the amounts that match HSE.",
        "voice": "Need the total overtime amount for HSE? Use SUMIFS. Excel adds only the rows where the department is HSE. The answer is seven hundred. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "countifs",
        "problem": "Need to count open HSE actions automatically?",
        "title": "Count records with COUNTIFS",
        "headers": ["Action", "Dept", "Status"],
        "rows": [["Action 01", "HSE", "Open"], ["Action 02", "HR", "Closed"], ["Action 03", "HSE", "Open"], ["Action 04", "HSE", "Closed"]],
        "formula": '=COUNTIFS(B2:B5,"HSE",C2:C5,"Open")',
        "result": "2",
        "explanation": "COUNTIFS counts rows matching both conditions.",
        "voice": "Need to count open HSE actions? Use COUNTIFS. Excel checks the department and status together, then counts the matching rows. The answer is two. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "iferror",
        "problem": "Want to stop ugly division errors in reports?",
        "title": "Clean errors with IFERROR",
        "headers": ["Sales", "Target", "Achievement"],
        "rows": [[10000, 10000, ""], [7500, 0, ""], [9000, 10000, ""]],
        "formula": '=IFERROR(A2/B2,0)',
        "result": "100%",
        "explanation": "IFERROR returns a safe value instead of an error.",
        "voice": "Want to keep a report clean when a calculation fails? Use IFERROR. Excel performs the calculation, and if it creates an error, it returns zero instead. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "networkdays",
        "problem": "Need to calculate working days between two dates?",
        "title": "Count working days with NETWORKDAYS",
        "headers": ["Start Date", "End Date", "Working Days"],
        "rows": [["01-Sep-2026", "14-Sep-2026", ""], ["07-Sep-2026", "11-Sep-2026", ""]],
        "formula": "=NETWORKDAYS(A2,B2)",
        "result": "10",
        "explanation": "NETWORKDAYS counts weekdays between two dates.",
        "voice": "Need to calculate working days between two dates? Use NETWORKDAYS. Excel counts the weekdays between the start and end dates. Here the result is ten working days. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "eomonth",
        "problem": "Need the last day of the month automatically?",
        "title": "Get month end with EOMONTH",
        "headers": ["Date", "Month End"],
        "rows": [["14-Sep-2026", ""], ["02-Oct-2026", ""], ["19-Nov-2026", ""]],
        "formula": "=EOMONTH(A2,0)",
        "result": "30-Sep-2026",
        "explanation": "EOMONTH returns the last day of the selected month.",
        "voice": "Need the last day of a month automatically? Use EOMONTH. Enter the date, use zero for the current month, and Excel returns the month end. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "textjoin",
        "problem": "Need to combine first and last names quickly?",
        "title": "Combine names with TEXTJOIN",
        "headers": ["First Name", "Last Name", "Full Name"],
        "rows": [["Ahmed", "Khan", ""], ["Sara", "Ali", ""], ["Omar", "Hassan", ""]],
        "formula": '=TEXTJOIN(" ",TRUE,A2:B2)',
        "result": "Ahmed Khan",
        "explanation": "TEXTJOIN combines cells with a chosen separator.",
        "voice": "Need to combine first and last names quickly? Use TEXTJOIN. Excel joins the cells with a space between them. No manual typing needed. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "trim",
        "problem": "Copied data has unwanted spaces?",
        "title": "Clean copied data with TRIM",
        "headers": ["Raw Name", "Clean Name"],
        "rows": [["  Ahmed   Khan  ", ""], [" Sara Ali ", ""], ["  Omar  ", ""]],
        "formula": "=TRIM(A2)",
        "result": "Ahmed Khan",
        "explanation": "TRIM removes extra spaces from text.",
        "voice": "Copied data full of unwanted spaces? Use TRIM. Excel removes extra spaces and gives you clean text. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "round",
        "problem": "Need clean numbers with two decimal places?",
        "title": "Round numbers with ROUND",
        "headers": ["Amount", "Rounded"],
        "rows": [[1254.678, ""], [987.456, ""], [42.995, ""]],
        "formula": "=ROUND(A2,2)",
        "result": "1254.68",
        "explanation": "ROUND keeps the number to two decimal places.",
        "voice": "Need clean numbers with two decimal places? Use ROUND. Enter the number and two as the number of decimals. Excel returns a clean rounded value. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "maxifs",
        "problem": "Need the highest sales value for one department?",
        "title": "Find the highest value with MAXIFS",
        "headers": ["Employee", "Dept", "Sales"],
        "rows": [["Ahmed", "HSE", 12000], ["Sara", "HR", 15000], ["Ali", "HSE", 18000], ["Omar", "HSE", 14000]],
        "formula": '=MAXIFS(C2:C5,B2:B5,"HSE")',
        "result": "18000",
        "explanation": "MAXIFS finds the largest value that meets a condition.",
        "voice": "Need the highest sales value for HSE? Use MAXIFS. Excel checks the department and returns the largest matching sales value. Here the result is eighteen thousand. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "unique",
        "problem": "Need a list without duplicate departments?",
        "title": "Remove duplicates with UNIQUE",
        "headers": ["Department", "Unique Departments"],
        "rows": [["HSE", ""], ["HR", ""], ["HSE", ""], ["Finance", ""]],
        "formula": "=UNIQUE(A2:A5)",
        "result": "HSE, HR, Finance",
        "explanation": "UNIQUE creates a list containing each value once.",
        "voice": "Need a list without duplicate departments? Use UNIQUE. Excel automatically creates a list containing each department once. Use it to clean a list quickly. Follow Learn Verse for more Excel tips."
    },
    {
        "id": "filter",
        "problem": "Need to show only open records?",
        "title": "Show matching rows with FILTER",
        "headers": ["Employee", "Status", "Department"],
        "rows": [["Ahmed", "Open", "HSE"], ["Sara", "Closed", "HR"], ["Ali", "Open", "Finance"]],
        "formula": '=FILTER(A2:C4,B2:B4="Open")',
        "result": "Ahmed, Ali",
        "explanation": "FILTER displays only rows that meet a condition.",
        "voice": "Need to show only open records? Use FILTER. Excel keeps the rows where status is Open and hides the rest from the result. Follow Learn Verse for more Excel tips."
    },
]

def get_daily_formulas(run_date: date, count=3):
    if count < 1 or count > len(FORMULAS):
        raise ValueError(f"count must be between 1 and {len(FORMULAS)}")
    days = (run_date - date(2026, 1, 1)).days
    offset = (days * 3) % len(FORMULAS)
    return [FORMULAS[(offset + i) % len(FORMULAS)] for i in range(count)]
