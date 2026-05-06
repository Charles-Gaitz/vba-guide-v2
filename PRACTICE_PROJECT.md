# VBA Practice for ACCT 628 — Practice Project Specification

## Aggie Advisors: New Student Onboarding

### The Scenario

You work in the advising office at Mays Business School. Each semester a new group
of students is accepted into the Professional Program in Accountancy (PPA). Your job
is to build a macro that processes the incoming applicant list, adds accepted students
to the master student roster, and updates the summary reports.

This scenario mirrors the Project Demo from class intentionally. The difference is
that you build it one module at a time — each module's Exam Challenge is one piece
of the complete macro. By the time you finish all seven modules, you have built the
whole thing yourself.

**No downloadable file is provided.** Copy the data tables below and paste them
into your own Excel workbook following the worksheet setup instructions.

---

## Workbook Setup Instructions

Create a new Excel file and save it as `Aggie_Advisors_Practice.xlsm`
(macro-enabled workbook). Create the following sheets in this order:

### Sheet 1 — Instructions
Leave this sheet for your own notes about the project.

### Sheet 2 — Applicant Information
- Create a table named `ApplicantData`
- Column headers (row 1): StudentID | LastName | FirstName | TAMU_GPR |
  Grade229 | Grade230 | Grade327 | FinalDecision
- Copy the 30 applicant records from the data table below into rows 2–31

### Sheet 3 — Student Information
- Create a table named `PPAData`
- Column headers (row 1): StudentID | Group | TrackCode | LastName | FirstName |
  UndergradGPR | GradGPR | Advisor | Employer1 | Employer2 | RecruitingStatus
- Copy the 85 existing student records from the data table below into rows 2–86

### Sheet 4 — Advisor Data
- Create a table named `AdvisorData`
- Column headers: TrackCode | TrackName | Advisor
- Enter these 4 rows:

| TrackCode | TrackName | Advisor |
|---|---|---|
| FI | Finance | Dr. Martinez |
| TX | Taxation | Dr. Patel |
| AC | Accounting | Dr. Johnson |
| U | Undecided | Dr. Williams |

### Sheet 5 — Group Summary
- Leave blank for now — you will add a pivot table in a later module
- Create a Named Range: select cell B2, go to Formulas → Name Manager → New,
  name it `CurrentGroup`, value will be set by the macro

### Sheet 6 — Valid Values
- Column A header: Group Number
- Enter groups 30 through 34 in rows 2–6 (these are the existing groups)
- Create a Named Range: select cell D2, name it `AcceptValue`, enter "Accept" as the value
- This is how the macro references "Accept" without hardcoding the string

---

## Named Ranges to Create (Formulas → Name Manager)

| Name | Points To | Initial Value |
|---|---|---|
| `AcceptValue` | Valid Values!D2 | Accept |
| `CurrentGroup` | Group Summary!B2 | 30 |

---

## Applicant Information — 30 Records (copy this table)

| StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision |
|---|---|---|---|---|---|---|---|
| 724816395 | Anderson | Emma | 3.842 | A | B+ | A- | Accept |
| 831924750 | Martinez | Carlos | 3.156 | B | B | B- | Deny |
| 619283740 | Thompson | Sarah | 3.971 | A | A- | A | Accept |
| 748291036 | Nguyen | Michael | 3.624 | B+ | A- | B+ | Accept |
| 825163947 | Williams | Ashley | 2.987 | C+ | B | C | Deny |
| 736481920 | Brown | James | 3.745 | A- | B+ | A- | Accept |
| 814729360 | Davis | Lauren | 3.512 | B+ | B+ | B | Accept |
| 729183640 | Wilson | Tyler | 2.843 | C | C+ | B- | Deny |
| 836492710 | Johnson | Megan | 3.889 | A | A | A- | Accept |
| 715824930 | Garcia | Daniel | 3.234 | B | B- | B | Deny |
| 842163950 | Miller | Rachel | 3.763 | A- | A- | B+ | Accept |
| 726849130 | Moore | Kevin | 3.091 | B- | C+ | B | Deny |
| 819374620 | Taylor | Jessica | 3.956 | A | A- | A | Accept |
| 734826190 | Jackson | Ryan | 3.478 | B+ | B | B+ | Accept |
| 821649370 | White | Brittany | 2.765 | C | B- | C+ | Deny |
| 716293840 | Harris | Brandon | 3.834 | A- | A- | A- | Accept |
| 843716290 | Martin | Stephanie | 3.612 | B+ | B+ | A- | Accept |
| 728493610 | Thompson | Nathan | 3.147 | B | B- | B | Deny |
| 815264930 | Lewis | Amanda | 3.891 | A | A | A | Accept |
| 731846920 | Robinson | Justin | 3.423 | B+ | B | B | Accept |
| 824619730 | Clark | Samantha | 2.914 | C+ | B- | C | Deny |
| 719283460 | Rodriguez | Matthew | 3.756 | A- | B+ | A- | Accept |
| 836142790 | Lee | Kayla | 3.534 | B+ | A- | B+ | Accept |
| 724891360 | Walker | Andrew | 3.068 | B- | C+ | B- | Deny |
| 811364920 | Hall | Courtney | 3.847 | A | A- | A | Accept |
| 738261490 | Allen | Christopher | 3.389 | B | B+ | B | Accept |
| 826419730 | Young | Melissa | 2.831 | C | C+ | B- | Deny |
| 713849260 | Hernandez | Joshua | 3.912 | A | A | A- | Accept |
| 841263970 | King | Tiffany | 3.645 | B+ | A- | B+ | Accept |
| 727491360 | Wright | Patrick | 3.178 | B | B- | B | Deny |

**Verification totals (use these to check your macro answers):**
- Total applicants: 30
- Accepted: 20
- Denied: 10
- Average GPR of accepted students: 3.7175
- Average GPR of denied students: 3.0014

---

## Student Information — 85 Existing Records

The 85 existing students represent Groups 30–34, all four track codes (FI, TX, AC, U),
and all employment fields filled in (not "Unknown" — these are existing students).

Generate these yourself or ask Claude to generate them using this structure.
The key constraints are:
- 85 records total in rows 2–86
- Groups 30–34 (approximately 17 per group)
- TrackCodes: FI (~25), TX (~20), AC (~22), U (~18)
- Advisor matches the AdvisorData table (Dr. Martinez for FI, etc.)
- UndergradGPR between 2.7 and 4.0
- GradGPR between 2.8 and 4.0
- Employer1, Employer2, RecruitingStatus: realistic firm names / statuses
  (not "Unknown" — these students have gone through recruiting)

When your Module 5 Exam Challenge macro runs, it should find row 87 as the
first empty row (86 rows of data + 1 header = next empty is row 87).

---

## The Complete Macro — Built One Module at a Time

Each module's Exam Challenge adds one section to this macro.
By Module 7 you have the complete working version.

```vba
Option Explicit
' ACCT 628 - Sanders
' Aggie Advisors: New Student Onboarding
' Module: Onboarding

Sub AddNewStudents()

    ' ── MODULE 2 ────────────────────────────────────────────────
    ' Speed settings
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' ── MODULE 3 ────────────────────────────────────────────────
    ' DEFINE Variables
    Dim NewGroup As Integer
    Dim NumberAccepted As Integer
    Dim StudentID As Long
    Dim GPR As Double
    Dim LastName As String
    Dim FirstName As String

    ' ── MODULE 2 ────────────────────────────────────────────────
    ' Check for records
    Sheets("Applicant Information").Select
    Range("A2").Select

    If ActiveCell = "" Then
        MsgBox "No applicants found. Please enter applicant data first."
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Prompt for group number
    NewGroup = InputBox("Enter the new group number:")

    ' ── MODULE 1 ────────────────────────────────────────────────
    ' Add NewGroup to Valid Values
    Sheets("Valid Values").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell = NewGroup
    ActiveCell.Offset(1, 0).Select

    ' Copy Applicant Information as backup
    Sheets("Applicant Information").Select
    Sheets("Applicant Information").Copy Before:=Sheets(6)
    ActiveSheet.Name = "Applicant Information Group " & NewGroup

    ' ── MODULE 5 ────────────────────────────────────────────────
    ' Navigate to first empty row on Student Information
    Sheets("Student Information").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("PPAData[[#Headers],[StudentID]]").Select

    ' Navigate to first applicant record
    Sheets("Applicant Information").Select
    Range("A2").Select

    ' ── MODULE 4 ────────────────────────────────────────────────
    ' Loop through all applicants
    Do Until ActiveCell = ""

        ' ── MODULE 6 ──────────────────────────────────────────
        ' Only process accepted students
        If ActiveCell.Offset(0, 7) = Range("AcceptValue") Then

            ' ── MODULE 3 ──────────────────────────────────────
            ' POPULATE variables
            StudentID = ActiveCell
            LastName  = ActiveCell.Offset(0, 1)
            FirstName = ActiveCell.Offset(0, 2)
            GPR       = ActiveCell.Offset(0, 3)

            ' ── MODULE 5 ──────────────────────────────────────
            ' SELECT Student Information and DISPLAY
            Sheets("Student Information").Select

            ActiveCell            = StudentID
            ActiveCell.Offset(0, 1) = NewGroup
            ActiveCell.Offset(0, 2) = "U"
            ActiveCell.Offset(0, 3) = LastName
            ActiveCell.Offset(0, 4) = FirstName
            ActiveCell.Offset(0, 5) = GPR
            ActiveCell.Offset(0, 6) = 0
            ActiveCell.Offset(0, 7) = "Dr. Williams"
            ActiveCell.Offset(0, 8) = "Unknown"
            ActiveCell.Offset(0, 9) = "Unknown"
            ActiveCell.Offset(0, 10) = "Unknown"

            NumberAccepted = NumberAccepted + 1

            ' MOVE to next empty student row
            ActiveCell.Offset(1, 0).Range("PPAData[[#Headers],[StudentID]]").Select

            ' RETURN to applicant list
            Sheets("Applicant Information").Select

        End If

        ' ALWAYS move to next applicant (accepted or denied)
        ActiveCell.Offset(1, 0).Range("ApplicantData[[#Headers],[StudentID]]").Select

    Loop

    ' ── MODULE 1 ────────────────────────────────────────────────
    ' Update Group Summary and refresh
    Sheets("Group Summary").Select
    Range("CurrentGroup") = NewGroup
    ActiveWorkbook.RefreshAll

    ' ── MODULE 2 ────────────────────────────────────────────────
    ' Restore settings and display completion message
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox NumberAccepted & " students added and reports updated for Group " & NewGroup

End Sub
```

---

## Module-by-Module Section Guide

### #module-1 — Macro Foundations
**What this piece does:** Navigates the workbook, bolds/autofits Student Information
headers, copies the Applicant sheet as a backup, adds the group number to Valid Values.

**Your task from Module 1:**
Record and modify a macro that:
1. Navigates to Student Information and bolds row 1 without SELECT/SELECTION
2. Autofits all columns on Student Information
3. Navigates back to Applicant Information
Expected: headers bolded, columns autofitted.
← [Back to Macro Foundations](/modules/foundations#exam-challenge)

---

### #module-2 — Adding Programming Concepts
**What this piece does:** Adds Option Explicit, variable declarations, the guard check
for empty applicant list, the InputBox for group number, speed settings on/off.

**Your task from Module 2:**
Write a macro that checks for records and prompts for group number.
Expected: displays "Group [X] is ready to process" — stops gracefully if no records.
← [Back to Adding Programming Concepts](/modules/programming-concepts#exam-challenge)

---

### #module-3 — Variables
**What this piece does:** Declares the variables that hold each student's data
(StudentID, LastName, FirstName, GPR) and populates them from each applicant row.

**Your task from Module 3:**
Declare correct data types and populate from Offset navigation.
Expected: MsgBox shows "ID: 724816395 | Name: Anderson, Emma | GPR: 3.842" for row 2.
← [Back to Variables](/modules/variables#exam-challenge)

---

### #module-4 — Loops
**What this piece does:** The Do Until loop that processes every applicant, the IF
that checks FinalDecision, and the NumberAccepted counter.

**Your task from Module 4:**
Loop through all 30 applicants, count accepted, calculate average GPR.
Expected: 20 accepted, average GPR 3.7175.
← [Back to Loops](/modules/loops#exam-challenge)

---

### #module-5 — Relative vs Absolute References
**What this piece does:** Navigating to the first empty row on Student Information
using End(xlDown) + Offset(1,0), and populating each field using Offset from ActiveCell.

**Your task from Module 5:**
Navigate to the correct empty row and display its row number.
Expected: row 87.
← [Back to Relative vs Absolute References](/modules/references#exam-challenge)

---

### #module-6 — Filters & Shortcut Keys
**What this piece does:** The IF statement inside the loop that checks
FinalDecision = Range("AcceptValue") before processing — the filter logic
that makes only accepted students get added.

**Your task from Module 6:**
Use AutoFilter approach to isolate accepted students as an alternative to the loop IF.
Expected: 20 rows copied to "Accepted Students" sheet, count matches Module 4.
← [Back to Filters & Shortcut Keys](/modules/filters#exam-challenge)

---

### #module-7 — F8 Debugging
**What this piece does:** This module gives you a broken version of the complete
AddNewStudents macro with three hidden bugs. Your job is to find and fix all three
using F8 and the Watch Window.

**The three bugs:**
1. Wrong column offset in the IF statement (checks column A instead of column H)
2. Move to next record inside the IF block instead of outside (causes endless loop on Deny)
3. Missing `NumberAccepted = 0` initialization (gives wrong count on second run)

**Your task from Module 7:**
Use F8 and Watch Window to find all three. Fix each with the minimum change.
Verify: corrected macro adds exactly 20 students, final MsgBox shows correct count.
← [Back to F8 Debugging](/modules/debugging#exam-challenge)

---

## Putting It All Together

Once you've completed all seven modules, you have built the complete AddNewStudents
macro. Run the full macro with group number 35. Expected result:
- 20 new students added to Student Information (rows 87–106)
- All new students: TrackCode = "U", Advisor = "Dr. Williams", GradGPR = 0
- Group Summary updated: CurrentGroup = 35
- MsgBox: "20 students added and reports updated for Group 35"

If your results differ, return to the module that covers the section that's wrong
and use F8 to debug it.
