"""Shared column layout for prompt experiment workbook."""

# Columns A–D (fixed template layout)
COL_SERIAL = 1
COL_SCENARIO = 2
COL_PROMPT_SCENARIO = 3
COL_LLM_RESPONSE = 4

RUBRIC_HEADERS = [
    "D1 Groundedness",
    "D2 Structural Completeness",
    "D3 Technical Depth",
    "D4 Uncertainty Handling",
    "D5 Actionability",
    "D6 Professional Quality",
    "Overall",
    "Notes",
]

COL_FIRST_SCORE = 5
COL_LAST_SCORE = 10
COL_OVERALL = 11
COL_NOTES = 12
FIRST_DATA_ROW = 2
LAST_DATA_ROW = 11
LAST_TEMPLATE_ROW = 21

PROMPT_SHEETS = ("Prompt A", "Prompt B", "Prompt C")
