# Excel Formula Shorts Automation

Automatically creates **3 vertical Excel formula Shorts per day** using GitHub Actions.

## What each video shows

Each 10–15 second video:
1. Opens an Excel-style worksheet.
2. Types realistic work data.
3. Selects the formula cell.
4. Types the Excel formula.
5. Calculates/shows the outcome.
6. Highlights the result.
7. Displays a short "Save this formula" ending.

The generator uses **Python + Pillow + FFmpeg**. Microsoft Excel is NOT required.

## Repository structure

```text
excel-formula-shorts/
├── .github/
│   └── workflows/
│       └── generate_excel_shorts.yml
├── output/
├── excel_formula_generator.py
├── formula_library.py
├── requirements.txt
├── .gitignore
└── README.md
```

## Daily schedule

The workflow runs once per day at **04:00 UTC = 08:00 UAE time**.

It generates 3 videos in one run. The selected formulas rotate automatically by date, so the same 3 formulas are not normally repeated on consecutive days.

You can also run it manually from:

**GitHub → Actions → Daily Excel Formula Shorts → Run workflow**

## Output

The three MP4 files are uploaded as a GitHub Actions artifact named:

`excel-formula-shorts`

Artifact retention is set to **3 days**.

No API key is required for video generation.

## Run locally

Install Python 3.11+ and FFmpeg, then:

```bash
pip install -r requirements.txt
python excel_formula_generator.py
```

The MP4 files will be created in `output/`.

## Important

The formula results in this project are generated from predefined, verified examples rather than by launching Microsoft Excel. This keeps GitHub Actions free of Office licensing/dependency issues.

You can expand `formula_library.py` with more formulas and workplace examples.
