# Relative vs Absolute References — Content Spec
# Module 6 of 9
# File: src/modules/references.html
# Prev: /src/modules/calculations.html (Calculations and Dates)
# Next: /src/modules/filters.html (Filters & Shortcut Keys)

---

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Relative vs Absolute References, you should
have already watched the Absolute vs Relative References Video in Canvas
and followed along with the Macro Demo file. This practice will build
upon that foundation."

---

## CONCEPT SECTION (id="concept")
h2 heading: "Relative vs Absolute References in VBA"

### Opening — Why References Matter

**Paragraph 1:**
When you record a macro, Excel defaults to absolute references —
every action is recorded as a specific cell address like Range("B3").
That means no matter where you are when you run the macro, it always
goes to B3. Relative references work differently: they record movement
relative to wherever the active cell is, using Offset. Understanding
when to use each one is one of the most practical skills in this course.

**Paragraph 2:**
The distinction matters most when you're building loops. Inside a loop,
you almost always want relative references — move down one row from
wherever you are. Outside a loop, navigating to a specific header or
starting cell, you want absolute. Mixing them up is one of the most
common sources of bugs in student macros.

**course-tip (concept):**
"If your macro keeps going to the wrong cell, the first thing to check
is whether you used absolute when you needed relative, or vice versa.
This is almost always the cause."

---

#### Absolute References

**Paragraph 1:**
Absolute references always go to the same cell regardless of where
the active cell is. They are the default when you record a macro.
Range("B3") always selects B3. This is what you want when navigating
to a fixed location — a header row, a named range, or a specific
starting cell.

**.syntax-box:**
```
' Always goes to cell B3 — absolute
Range("B3").Select

' Always populates A1 — absolute
Range("A1") = "Sales Report"

' Navigate to the start of data — absolute
Range("A2").Select
```

---

#### Relative References

**Paragraph 1:**
Relative references move relative to the current active cell using
Offset. Offset(rows, columns) — positive rows move down, positive
columns move right, negative values move up or left. This is what
you want inside loops when you need to process the current row and
then move to the next one.

**.syntax-box:**
```
' Move down one row from wherever you are
ActiveCell.Offset(1, 0).Select

' Move one column to the right
ActiveCell.Offset(0, 1).Select

' Move two rows down, one column left
ActiveCell.Offset(2, -1).Select

' Populate the cell one row below without selecting it
ActiveCell.Offset(1, 0) = "Done"
```

**Introduction sentence before code:**
"Inside a loop, relative references are essential — this pattern
moves through every row without hardcoding any cell address:"

```vba
Range("A2").Select

Do Until ActiveCell = ""
    ' process current row using Offset
    ActiveCell.Offset(0, 1) = "Processed"

    ' always move down one row at end of loop
    ActiveCell.Offset(1, 0).Select
Loop
```

---

#### Switching Between the Two

**Paragraph 1:**
You switch to relative recording in Excel via the Developer ribbon →
Use Relative References button. When active, recorded actions use
Offset instead of Range addresses. Toggle it off to go back to absolute.
In code you write yourself, you choose by deciding whether to use
Range("address") or ActiveCell.Offset(row, col).

**.syntax-box:**
```
' Recorded in absolute mode
Range("B3").Select
Selection.Font.Bold = True

' Same action recorded in relative mode
ActiveCell.Offset(0, 1).Select
Selection.Font.Bold = True

' Written directly — populate without selecting
Range("B3").Font.Bold = True          ' absolute, no Select needed
ActiveCell.Offset(0, 1).Font.Bold = True   ' relative, no Select needed
```

---

#### Named Ranges and Why They Matter

**Paragraph 1:**
Hardcoded cell references like Range("B3") break if someone inserts
a row above row 3 — your code still goes to B3 but that's now the
wrong cell. Named Ranges solve this. A Named Range always follows
the data it was assigned to, regardless of where rows or columns move.
In VBA, reference them exactly like any range: Range("MyNamedRange").

**.syntax-box:**
```
' Hardcoded — breaks if rows are inserted above
Range("B3") = NewGroup

' Named Range — always finds the right cell
Range("CurrentGroup") = NewGroup

' Named Range in a formula
Range("AverageRange") = "=TotalRange/CountRange"
```

---

### QUICK CHECK SECTION (id="quick-check")

**Question 1:**
What is the default reference type when you record a macro in Excel?
- A. Relative
- B. Absolute ← CORRECT
- C. Mixed
- D. It depends on which cell you start on
**Explanation:** Excel records in absolute mode by default. Every action
is recorded as a specific cell address. Switch to relative mode using
Developer → Use Relative References before recording.

**Question 2:**
Your ActiveCell is B5. What cell does `ActiveCell.Offset(2, -1)` reference?
- A. B7
- B. A7 ← CORRECT
- C. C3
- D. A3
**Explanation:** Offset(2, -1) means 2 rows down and 1 column left.
From B5: 2 rows down = row 7, 1 column left = column A. Result: A7.

**Question 3:**
Inside a loop that processes rows of data, which reference type should
you use to move to the next row?
- A. Absolute — Range("A" & rowNumber).Select
- B. Relative — ActiveCell.Offset(1, 0).Select ← CORRECT
- C. Either works the same way
- D. Neither — loops move automatically
**Explanation:** Inside a loop you're at a different row each iteration.
Relative references move from wherever you are — that's exactly what
you need. Absolute would always go back to the same row.

**Question 4:**
You insert a new row at the top of your data. Your code has
`Range("A2").Select` as the starting point. What happens?
- A. The code finds the new first data row automatically
- B. The code still selects A2, which is now the wrong row ← CORRECT
- C. VBA throws an error
- D. The named range updates automatically
**Explanation:** Absolute cell references in code don't adjust when
rows are inserted. Range("A2") always goes to A2. Use Named Ranges
for starting points that might shift.

**Question 5:**
What is the advantage of using a Named Range over a hardcoded cell address?
- A. Named Ranges run faster
- B. Named Ranges work in formulas but hardcoded addresses don't
- C. Named Ranges follow the data if rows or columns are inserted ← CORRECT
- D. There is no advantage — they work identically
**Explanation:** Named Ranges are tied to the data, not the address.
If rows shift, the Named Range still points to the right cell.
Hardcoded addresses always go to the same address regardless.

**course-tip after quick check:**
"Question 2 is the exact type you'll see on the exam — given a starting
cell and an Offset, find the target. Practice these until they're instant."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Predict the Cell (SIMPLE FORMAT)
**Difficulty:** Observation

Without running any code, predict where each Offset lands if
ActiveCell is currently D10. Write your predictions down, then
paste this macro into VBA and run it to verify:

```vba
Option Explicit
Sub OffsetPractice()
    Sheets("Sheet1").Select
    Range("D10").Select

    MsgBox "Offset(1,0): " & ActiveCell.Offset(1, 0).Address
    MsgBox "Offset(0,1): " & ActiveCell.Offset(0, 1).Address
    MsgBox "Offset(-2,3): " & ActiveCell.Offset(-2, 3).Address
    MsgBox "Offset(0,-3): " & ActiveCell.Offset(0, -3).Address
End Sub
```
Each MsgBox pauses and shows the target cell address. Press OK to see the next one.

**Hint:** Address property returns the cell address as a string like "$E$10".

**Solution:**
- Offset(1, 0) → D11 (one row down)
- Offset(0, 1) → E10 (one column right)
- Offset(-2, 3) → G8 (two rows up, three columns right)
- Offset(0, -3) → A10 (three columns left)

---

#### Exercise 2 — Record Both Ways (STEPS FORMAT)
**Difficulty:** Guided

Record the same macro twice — once in absolute mode, once in relative
mode — and compare the code.

**Step 1 — Record in absolute mode**
Go to Developer → Record Macro. Name it AbsoluteTest.
Click cell C5. Type "Hello". Press Enter. Stop recording.
Open VBA Editor and note what the code looks like.

**Step 2 — Record in relative mode**
Go to Developer → Use Relative References (turn it ON).
Go to Developer → Record Macro. Name it RelativeTest.
Click cell C5. Type "Hello". Press Enter. Stop recording.
Turn off Use Relative References.

**Step 3 — Compare the two macros**
Open both macros in the VBA Editor side by side. You should see:

AbsoluteTest:
```vba
Range("C5").Select
ActiveCell.FormulaR1C1 = "Hello"
```
This always goes to C5 no matter what cell was selected before running.

RelativeTest:
```vba
ActiveCell.FormulaR1C1 = "Hello"
```
This types "Hello" wherever the active cell is when you run it.

**Solution:**
Test it: click cell A1, run RelativeTest — "Hello" appears in A1.
Click cell F8, run RelativeTest — "Hello" appears in F8.
Click cell A1, run AbsoluteTest — "Hello" always appears in C5.
The key difference: AbsoluteTest has a hardcoded Range("C5"), RelativeTest
has no cell address at all — it acts on wherever the cursor is.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table (Aggie Advisors — 30 records)
Full 30-record dataset from PRACTICE_PROJECT.md.
Columns: StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision

#### Practice Problem — Navigate to the First Empty Row

**Setup (do this before writing the macro):**
1. Create a new sheet in your workbook named "Student Information"
2. Add these column headers in row 1:
   StudentID | Group | TrackCode | LastName | FirstName | UndergradGPR
3. Copy the first 10 rows of student data from the Aggie Advisors table
   below and paste them starting at row 2 (so rows 2-11 have data,
   row 12 is empty)

**What your macro needs to do:**
- Navigate to the Student Information sheet using an absolute reference
- Go to cell A2 (absolute — always start here)
- Use Selection.End(xlDown).Select to jump to the last filled row (row 11)
- Use ActiveCell.Offset(1, 0).Select to move to the first empty row (row 12)
- Display: "Ready to add student at row " & ActiveCell.Row

**Expected result:** "Ready to add student at row 12"

**Hint:**
```vba
Sheets("Student Information").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
MsgBox "Ready to add student at row " & ActiveCell.Row
```

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-6"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Relative vs Absolute References in VBA
**No hints. Exam level.**

**Setup:** Use the same Student Information sheet from the practice problem
(10 rows of data in rows 2-11, first empty row is row 12). Keep the
Aggie Advisors Applicant Information data on Sheet1.

Write a macro that:
1. Navigates to Sheet1 (Applicant Information), goes to cell A2
2. Reads the first applicant record into variables:
   StudentID (Long) from column A,
   LastName (String) from column B (Offset 0,1),
   FirstName (String) from column C (Offset 0,2),
   GPR (Double) from column D (Offset 0,3)
3. Navigates to the Student Information sheet
4. Goes to cell A2, uses End(xlDown) to find the last row,
   uses Offset(1,0) to move to the first empty row
5. Populates that row using Offset from ActiveCell — no hardcoded addresses:
   ActiveCell = StudentID
   ActiveCell.Offset(0,1) = "35" (group number)
   ActiveCell.Offset(0,2) = "U" (track code)
   ActiveCell.Offset(0,3) = LastName
   ActiveCell.Offset(0,4) = FirstName
   ActiveCell.Offset(0,5) = GPR

Must use Option Explicit with all variables declared.
Must use absolute reference only for the A2 starting navigation.
Must use relative references (Offset) for all population steps.

**Expected result:** Row 12 of Student Information contains:
724816395 | 35 | U | Anderson | Emma | 3.842

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-6"`