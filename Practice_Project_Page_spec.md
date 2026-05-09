# Practice Project Page — Content Spec
# File: src/modules/practice-project.html
# This is NOT a teaching module — it is the Aggie Advisors scenario hub
# Prev: /src/modules/pseudocode.html (Pseudocode)
# Next: none (last page)

---

## PAGE STRUCTURE NOTE

This page is different from all other modules. It has NO:
- Quick Check section
- Easy Wins section
- .exercise-steps or .exercise-simple components
- .pseudocode-block or .code-block in the main content
  (except the complete macro in the "Putting It All Together" section)

It DOES have:
- .page-intro with title and subtitle
- .box-reminder
- In-page anchor nav (different from module nav — see below)
- Scenario overview section
- Workbook setup instructions
- Copyable data table (Applicant Information — 30 records)
- Module-by-module sections (#module-1 through #module-9)
- "Putting It All Together" section with complete macro code block
- .module-nav at bottom

---

## Page metadata

- Title: 'Practice Project — VBA Practice for ACCT 628'
- Meta description: 'The Aggie Advisors practice project for ACCT 628 —
  build the complete AddNewStudents macro one module at a time using
  the full Aggie Advisors scenario.'
- .page-title: 'Practice Project'
- .page-subtitle: 'Aggie Advisors: New Student Onboarding'

---

## .box-reminder

"This practice project runs throughout all 9 modules. Each module's Exam
Challenge is one piece of the complete AddNewStudents macro. Work through
the modules in order — by the time you finish, you will have built the
entire macro yourself."

---

## Anchor Nav (inside header — 6 links)

The anchor nav on this page has 6 links (different from the standard 5):
- href="#scenario" text "Scenario"
- href="#setup" text "Workbook Setup"
- href="#data" text "Data Table"
- href="#modules" text "Module Guide"
- href="#complete-macro" text "Complete Macro"
- href="#final-run" text "Final Run"

---

## SCENARIO SECTION (id="scenario")

h2: "The Aggie Advisors Scenario"

**Paragraph 1 (exact from PRACTICE_PROJECT.md):**
You work in the advising office at Mays Business School. Each semester a new group
of students is accepted into the Professional Program in Accountancy (PPA). Your job
is to build a macro that processes the incoming applicant list, adds accepted students
to the master student roster, and updates the summary reports.

**Paragraph 2 (exact from PRACTICE_PROJECT.md):**
This scenario mirrors the Project Demo from class intentionally. The difference is
that you build it one module at a time — each module's Exam Challenge is one piece
of the complete macro. By the time you finish all nine modules, you have built the
whole thing yourself.

**Paragraph 3:**
No downloadable file is provided. Copy the data tables below and paste them
into your own Excel workbook following the worksheet setup instructions.

---

## WORKBOOK SETUP SECTION (id="setup")

h2: "Workbook Setup Instructions"

Intro line: "Create a new Excel file and save it as Aggie_Advisors_Practice.xlsm
(macro-enabled workbook). Create the following sheets in this order:"

Render as 6 subsections with h3 headings:

**h3: Sheet 1 — Instructions**
Leave this sheet for your own notes about the project.

**h3: Sheet 2 — Applicant Information**
- Table name: ApplicantData
- Column headers (row 1): StudentID | LastName | FirstName | TAMU_GPR |
  Grade229 | Grade230 | Grade327 | FinalDecision
- Copy the 30 applicant records from the data table below into rows 2–31

**h3: Sheet 3 — Student Information**
- Table name: PPAData
- Column headers (row 1): StudentID | Group | TrackCode | LastName | FirstName |
  UndergradGPR | GradGPR | Advisor | Employer1 | Employer2 | RecruitingStatus
- This sheet will have 85 existing student records when fully set up.
  For exercises, you can start with just the header row and add records as needed.

**h3: Sheet 4 — Advisor Data**
- Table name: AdvisorData
- Column headers: TrackCode | TrackName | Advisor
- Enter these 4 rows as an HTML table:

| TrackCode | TrackName | Advisor |
| FI | Finance | Dr. Martinez |
| TX | Taxation | Dr. Patel |
| AC | Accounting | Dr. Johnson |
| U | Undecided | Dr. Williams |

**h3: Sheet 5 — Group Summary**
- Leave blank for now
- Create a Named Range: select cell B2, name it CurrentGroup

**h3: Sheet 6 — Valid Values**
- Column A header: Group Number
- Enter groups 30 through 34 in rows 2–6
- Create a Named Range: select cell D2, name it AcceptValue, value = "Accept"

After the sheet instructions, render the Named Ranges as a small table:
| Name | Points To | Initial Value |
| AcceptValue | Valid Values!D2 | Accept |
| CurrentGroup | Group Summary!B2 | 30 |

---

## DATA TABLE SECTION (id="data")

h2: "Applicant Information Data"

Instructions paragraph:
"Copy this data and paste it into cell A1 of your Applicant Information sheet.
Press Ctrl+V — Excel will split the columns automatically. Then format the
data as a table named ApplicantData."

.data-table-section:
- .data-table-toggle button: "📋 Show Data Table"
- .data-table-wrap hidden:
  - .copy-data-btn with full TSV data from PRACTICE_PROJECT.md
    (header row + all 30 records, tabs between columns, newlines between rows)
  - HTML table with thead (8 columns) and tbody (30 rows)
    Use the exact 30 records from PRACTICE_PROJECT.md

Verification totals after the data table section (plain paragraph):
"Verification totals — use these to check your macro answers:
Total applicants: 30 | Accepted: 20 | Denied: 10 |
Average GPR accepted: 3.7175 | Average GPR denied: 3.0014"

---

## MODULE GUIDE SECTION (id="modules")

h2: "Module-by-Module Guide"

Intro paragraph:
"Each module's Exam Challenge adds one piece to the complete macro.
Click any module below to jump to its section, or go directly to the
module page to review the concept."

Render each module as a subsection with its own anchor ID.
Each subsection follows this structure:
- h3 with the module name
- "What this piece does:" bold — one sentence description
- "Your task:" bold — the specific task
- Expected result line
- Back link to the module's exam challenge page

---

### #module-1 — Macro Foundations

h3: "Module 1 — Macro Foundations"

**What this piece does:** Navigates the workbook, copies the Applicant Information
sheet as a backup, and adds the new group number to Valid Values.

**Your task:**
Record and modify a macro that:
1. Navigates to Student Information and bolds row 1 without SELECT/SELECTION
2. Autofits all columns on Student Information
3. Navigates back to Applicant Information

**Expected result:** Headers bolded, columns autofitted.

Back link: href="/src/modules/foundations.html#challenge" text "← Back to Macro Foundations"

---

### #module-2 — Adding Programming Concepts

h3: "Module 2 — Adding Programming Concepts"

**What this piece does:** Adds Option Explicit, variable declarations, the guard
check for empty applicant list, the InputBox for group number, and speed settings.

**Your task:**
Write a macro that checks for records and prompts for group number.
Expected: displays "Group [X] is ready to process" — stops gracefully if no records.

Back link: href="/src/modules/programming-concepts.html#challenge" text "← Back to Adding Programming Concepts"

---

### #module-3 — Variables

h3: "Module 3 — Variables"

**What this piece does:** Declares the variables that hold each student's data
(StudentID, LastName, FirstName, GPR) and populates them from each applicant row.

**Your task:**
Declare correct data types and populate from Offset navigation.
Expected: MsgBox shows "ID: 724816395 | Name: Anderson, Emma | GPR: 3.842" for row 2.

Back link: href="/src/modules/variables.html#challenge" text "← Back to Variables"

---

### #module-4 — Loops

h3: "Module 4 — Loops"

**What this piece does:** The Do Until loop that processes every applicant, the IF
that checks FinalDecision, and the NumberAccepted counter.

**Your task:**
Loop through all 30 applicants, count accepted, calculate average GPR.
Expected: 20 accepted, average GPR 3.7175.

Back link: href="/src/modules/loops.html#challenge" text "← Back to Loops"

---

### #module-5 — Calculations and Dates

h3: "Module 5 — Calculations and Dates"

**What this piece does:** Accumulates GPR totals, calculates averages using
WorksheetFunction, and formats results for display.

**Your task:**
Loop through all 30 records, accumulate total GPR for accepted students,
calculate average after the loop using Format(value, "0.000").
Expected: Accepted: 20 | Avg GPR: 3.718

Back link: href="/src/modules/calculations.html#challenge" text "← Back to Calculations and Dates"

---

### #module-6 — Relative vs Absolute References

h3: "Module 6 — Relative vs Absolute References"

**What this piece does:** Navigates to the first empty row on Student Information
using End(xlDown) + Offset(1,0), and populates each field using Offset from ActiveCell.

**Your task:**
Navigate to the correct empty row using absolute reference to A2,
then End(xlDown) and Offset(1,0) to find the first empty row.
Expected: row 87 (with 85 existing records + header).

Back link: href="/src/modules/references.html#challenge" text "← Back to Relative vs Absolute References"

---

### #module-7 — Filters & Shortcut Keys

h3: "Module 7 — Filters & Shortcut Keys"

**What this piece does:** The IF statement inside the loop that checks
FinalDecision = Range("AcceptValue") before processing — plus the Option 2
AutoFilter approach for isolating accepted students.

**Your task:**
Use AutoFilter to isolate accepted students and copy them to a new sheet.
Expected: 20 rows on "Accepted Students" sheet, count matches Module 4.

Back link: href="/src/modules/filters.html#challenge" text "← Back to Filters &amp; Shortcut Keys"

---

### #module-8 — F8 Debugging Practice

h3: "Module 8 — F8 Debugging Practice"

**What this piece does:** A broken version of the complete AddNewStudents macro.
Use F8 and the Watch Window to find and fix all bugs.

**The bugs to find:**
1. Wrong column offset in the IF statement (checks column A instead of column H)
2. Move to next record inside the IF block instead of outside
3. Missing NumberAccepted = 0 initialization

**Your task:**
Use F8 and Watch Window to find all three bugs.
Fix each with the minimum change — one line per bug.
Expected: corrected macro adds exactly 20 students.

Back link: href="/src/modules/debugging.html#challenge" text "← Back to F8 Debugging Practice"

---

### #module-9 — Pseudocode

h3: "Module 9 — Pseudocode"

**What this piece does:** Write complete pseudocode for the full AddNewStudents
macro using Sanders format — this is the planning document you would write
BEFORE building the macro.

**Your task:**
Write pseudocode covering all 7 sections: variable definitions, guard check,
user prompt, backup copy, main loop, refresh reports, completion message.
Your pseudocode should be detailed enough that someone who doesn't know VBA
could understand exactly what the macro does.

Back link: href="/src/modules/pseudocode.html#challenge" text "← Back to Pseudocode"

---

## COMPLETE MACRO SECTION (id="complete-macro")

h2: "The Complete Macro"

Intro paragraph:
"Once you have completed all nine modules, here is the complete AddNewStudents
macro you will have built. Each section is labeled with which module introduced it."

Render the complete macro as a .code-block with the full VBA from PRACTICE_PROJECT.md.
This is the ONLY .code-block on this page.
Include the full macro exactly as written in PRACTICE_PROJECT.md with all
module annotation comments (MODULE 1, MODULE 2, etc.)
Escape all & as &amp; in the code block.

---

## FINAL RUN SECTION (id="final-run")

h2: "Putting It All Together"

**Paragraph 1 (from PRACTICE_PROJECT.md):**
Once you've completed all nine modules, you have built the complete AddNewStudents
macro. Run the full macro with group number 35.

Expected result (render as a bullet list):
- 20 new students added to Student Information (rows 87–106)
- All new students: TrackCode = "U", Advisor = "Dr. Williams", GradGPR = 0
- Group Summary updated: CurrentGroup = 35
- MsgBox: "20 students added and reports updated for Group 35"

**Paragraph 2:**
If your results differ, return to the module that covers the section that's wrong
and use F8 to debug it. The module guide above links directly to each module's
Exam Challenge.

---

## MODULE NAV

prev: href="/src/modules/pseudocode.html" text "← Pseudocode"
center: href="/" text "All Modules"
next: empty span (this is the last page — no next)

For the next slot, render an empty non-linked span so layout stays balanced:
<span aria-hidden="true"></span>