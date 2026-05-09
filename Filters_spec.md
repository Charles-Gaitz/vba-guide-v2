# Filters & Shortcut Keys — Content Spec
# Module 7 of 9
# File: src/modules/filters.html
# Prev: /src/modules/references.html (Relative vs Absolute References)
# Next: /src/modules/debugging.html (F8 Debugging Practice)

---

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Filters & Shortcut Keys, you should have
already watched the Filters and Shortcut Keys (Option 2) Video in Canvas
and followed along with the Macro Demo file. This practice will build
upon that foundation."

---

## CONCEPT SECTION (id="concept")
h2 heading: "Filters and Shortcut Keys in VBA"

### Opening — Two Ways to Process Data

**Paragraph 1:**
So far you've processed data by looping through every record and using
an IF statement to act on the ones you want. That's Option 1 — it's
readable, easy to debug with F8, and works for any condition. Option 2
takes a different approach: filter the data to only the records you want,
then copy or process all of them at once. Both approaches produce the
same result. Option 2 is faster for large datasets. Option 1 is easier
to follow and debug. Knowing both gives you flexibility.

**Paragraph 2:**
This module covers the Option 2 techniques: AutoFilter to show only
matching rows, the FILTER function to extract records into a new location,
PasteSpecial to break formula links, and CountA to count results without
hardcoding numbers. These are the tools from the Filters and Shortcut
Keys demo video.

**course-tip (concept):**
"Option 1 is what the Step-Through and Build videos focus on. Option 2
shows up in the Filters demo. The exam may ask you to identify which
approach a given piece of code is using — know the signature patterns
of each."

---

#### AutoFilter

**Paragraph 1:**
AutoFilter shows only the rows that match a condition and hides the rest.
In VBA you apply it to a table using ListObjects and specify which field
(column number) and which criteria to filter on. After filtering, only
matching rows are visible — you can then copy, delete, or count them.
Always remove the filter when you're done.

**.syntax-box:**
```
' Apply AutoFilter — Field is the column number in the table
ActiveSheet.ListObjects("TableName").Range.AutoFilter Field:=8, Criteria1:="Deny"

' Delete visible (filtered) rows
ActiveCell.Rows("1:1").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp

' Remove the filter
ActiveSheet.ListObjects("TableName").Range.AutoFilter Field:=8
```

**Introduction sentence before code:**
"This is the Option 2 pattern from the Project Demo — filtering for
denied students and deleting them:"

```vba
' FILTER for Denied Students
ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8, _
    Criteria1:="Deny"

' DELETE Denied Students
ActiveCell.Rows("1:1").EntireRow.Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp

' Remove filter
ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8
```

---

#### FILTER Function in a Cell

**Paragraph 1:**
The FILTER function extracts matching records from a table into a new
location as a formula. In VBA you write it as a string into a cell using
Formula2R1C1. The result is a dynamic array — it spills into as many
rows as there are matching records. Because it's a formula, you usually
follow it with a copy/PasteSpecial step to convert the results to values
so they're no longer linked to the original data.

**.syntax-box:**
```
' Write FILTER formula into a cell
Range("A3").Formula2R1C1 = "=FILTER(TableName, TableName[Field]=""Value"",)"

' Copy the results
Range("A3").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

' PasteSpecial Values — breaks formula link
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
```

---

#### PasteSpecial Values

**Paragraph 1:**
When you copy cells that contain formulas and paste them normally,
the pasted cells still contain formulas linked to the original data.
PasteSpecial with xlPasteValues pastes only the results — plain numbers
and text with no formulas. This is essential after using FILTER or any
other formula-based extraction, because you want the data to stand on
its own without depending on the source table.

**.syntax-box:**
```
Selection.Copy

Selection.PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, _
    SkipBlanks:=False, _
    Transpose:=False
```

---

#### CountA to Count Results

**Paragraph 1:**
After filtering or copying data, you often need to know how many records
you ended up with. CountA counts non-blank cells — use it on a full column
after your filter or paste operation to get the count without hardcoding
a number. Subtract 1 if your range includes the header row.

**.syntax-box:**
```
' Count non-blank cells in column A (includes header)
NumberAccepted = Application.WorksheetFunction.CountA(Range("A:A")) - 1

' Or count just the data rows
NumberAccepted = Application.WorksheetFunction.CountA(Range("A2:A1000"))
```

---

### QUICK CHECK SECTION (id="quick-check")

**Question 1:**
What is the main difference between Option 1 (Loop + IF) and
Option 2 (Filter + Copy)?
- A. Option 1 only works on small datasets
- B. Option 2 processes one record at a time
- C. Option 1 loops through records individually; Option 2 filters first then acts on all matches at once ← CORRECT
- D. They produce different results
**Explanation:** Both produce the same results. Option 1 reads each
record, checks a condition, and acts if true. Option 2 filters to only
matching records first, then copies or processes all of them together.

**Question 2:**
After applying AutoFilter, you want to delete all visible rows.
What must you do after the deletion?
- A. Nothing — AutoFilter turns off automatically
- B. Run RefreshAll
- C. Remove the AutoFilter ← CORRECT
- D. Save the file
**Explanation:** AutoFilter stays on until you explicitly remove it.
Always turn it off after you're done so the table returns to showing
all records.

**Question 3:**
Why do you use PasteSpecial Values after copying FILTER results?
- A. PasteSpecial is faster than regular paste
- B. To break the formula link so the data stands alone ← CORRECT
- C. Regular paste doesn't work after AutoFilter
- D. To convert dates to text
**Explanation:** FILTER results are formulas linked to the source table.
PasteSpecial Values replaces the formulas with their current values —
plain data that doesn't change if the source changes.

**Question 4:**
You use CountA on column A after pasting filtered data. Your data has
30 records plus a header row. What does CountA(Range("A:A")) return?
- A. 30
- B. 31 ← CORRECT
- C. 29
- D. It depends on the filter
**Explanation:** CountA counts all non-blank cells including the header.
With 30 data rows and 1 header, it returns 31. Subtract 1 to get the
record count.

**Question 5:**
Which approach is easier to debug using F8?
- A. Option 2 (Filter + Copy) because it's fewer lines
- B. Option 1 (Loop + IF) because you can step through each record ← CORRECT
- C. They are equally easy to debug
- D. Neither can be debugged with F8
**Explanation:** Option 1 processes one record at a time so you can
watch each decision happen with F8. Option 2 operates on all matching
records at once — harder to inspect mid-execution.

**course-tip after quick check:**
"If an exam question shows you a macro and asks what it does, look for
the signature patterns: a Do Until loop with IF = Option 1, AutoFilter
or FILTER function = Option 2."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Apply and Remove AutoFilter (STEPS FORMAT)
**Difficulty:** Guided

Apply AutoFilter to the Aggie Advisors data to show only accepted
students, then count and remove the filter.

**Step 1 — Paste in your data**
Use the Aggie Advisors data from the data table below.
Make sure it's in a table named ApplicantData on Sheet1.
If your table has a different name, update the code accordingly.

**Step 2 — Write the filter macro**
In the VBA Editor, create a new Sub and add:
```vba
Sheets("Sheet1").Select
ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8, _
    Criteria1:="Accept"
```
Field:=8 targets column 8 of the table (FinalDecision).
Run it — you should see only the 20 accepted students.

**Step 3 — Count the visible rows**
Add this line after the AutoFilter:
```vba
Dim AcceptCount As Integer
AcceptCount = Application.WorksheetFunction.CountA(Range("A:A")) - 1
MsgBox "Accepted students: " & AcceptCount
```
CountA counts all non-blank cells including the header, so subtract 1.

**Step 4 — Remove the filter**
Add this line last:
```vba
ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8
```
Run the full macro. Filter applies, count shows 20, filter removes.

**Complete Code (View Complete Code):**
```vba
Option Explicit
Sub FilterAndCount()
    Dim AcceptCount As Integer

    Sheets("Sheet1").Select

    ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8, _
        Criteria1:="Accept"

    AcceptCount = Application.WorksheetFunction.CountA(Range("A:A")) - 1
    MsgBox "Accepted students: " & AcceptCount

    ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8
End Sub
```
**Expected result:** MsgBox shows "Accepted students: 20"

---

#### Exercise 2 — CountA Observation (SIMPLE FORMAT)
**Difficulty:** Observation

Run this standalone macro on your Aggie Advisors data.
It filters for accepted students, counts the visible rows,
displays the count, then removes the filter — all in one pass.
Before running, predict: what number will the MsgBox show?

```vba
Option Explicit
Sub CountAccepted()
    Dim AcceptCount As Integer

    Sheets("Sheet1").Select

    ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8, _
        Criteria1:="Accept"

    AcceptCount = Application.WorksheetFunction.CountA(Range("A:A")) - 1
    MsgBox "Accepted students: " & AcceptCount

    ActiveSheet.ListObjects("ApplicantData").Range.AutoFilter Field:=8
End Sub
```

**Hint:** CountA(Range("A:A")) counts every non-blank cell in column A,
including the header row. The filter hides non-matching rows but they
still exist — CountA only counts what's visible when a filter is active.

**Solution:** MsgBox shows "Accepted students: 20".
CountA sees 21 non-blank cells (20 accepted rows + 1 header), minus 1 = 20.
The 10 denied rows are hidden by the filter so CountA skips them.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table (Aggie Advisors — 30 records)
Full 30-record dataset from PRACTICE_PROJECT.md.
Columns: StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision

#### Practice Problem — Filter and Copy Accepted Students
Using the Aggie Advisors data, write a macro that uses the Option 2
approach to isolate accepted students and copy them to a new sheet.

**What your macro needs to do:**
- Navigate to Sheet1 where ApplicantData table lives
- Apply AutoFilter to ApplicantData, Field 8, Criteria1:="Accept"
- Select the visible data rows (not the header) and copy them
- Create or navigate to a sheet named "Accepted Students"
- PasteSpecial Values to paste without formula links
- Navigate back to Sheet1 and remove the AutoFilter
- Count the pasted records on "Accepted Students" using CountA
  and display: "Accepted students copied: [X]"

**Expected result:** 20 rows on the Accepted Students sheet.
MsgBox shows "Accepted students copied: 20"

**Hint — copying visible filtered rows:**
After applying AutoFilter, use this pattern to select and copy
only the visible data rows (skipping the header):
```vba
' Select visible rows after filter — skip header
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

' Navigate to destination and PasteSpecial
Sheets("Accepted Students").Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
```

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-7"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Compare Both Approaches
**No hints. Exam level.**

Write two macros that produce the same result using different approaches:

**Macro 1 — Option1_Count:** Use a Do Until loop with an IF statement
to count accepted students. Display the count in a MsgBox.

**Macro 2 — Option2_Count:** Use AutoFilter to filter for accepted
students, use CountA to count them, remove the filter, and display
the count in a MsgBox.

Both macros must display the same number.

After writing both, add a comment at the top of each macro explaining
in one sentence when you would choose that approach over the other.

**Expected result:** Both MsgBoxes show 20.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-7"`