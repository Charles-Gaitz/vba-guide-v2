# VBA Practice for ACCT 628 — Content Specification

## Ground Rules (apply to every module)

- All VBA code examples come directly from this file. Never invent examples.
- All terminology matches Professor Sanders exactly:
  - "DEFINE, POPULATE, USE" — the three steps of variables
  - "recorded" vs "typed" — code types
  - "hardcoded" — values that should be replaced with variables
  - "ActiveCell", "Offset", "Do Until" — primary navigation and loop patterns
- Pseudocode: keywords CAPITALIZED, indented inside loops/IF, END IF, ENDLOOP
- The Sanders approach is always the reference. AI code appears ONLY in .ai-compare.
- .course-tip voice: peer-to-peer, CE to fellow students, short and direct.
- Concept prose: 1–2 tight paragraphs per concept. No filler. No restatements.
  Every sentence must add new information.
- Each concept follows this structure:
  prose → .syntax-box (skeleton) → intro sentence → .code-block (full example)
- Exercise format per module is specified explicitly — steps or simple.

---

## Module 4 — Loops

### Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Loops, you should have already watched the Loops
Video in Canvas and followed along with the Macro Demo file.
This practice will build upon that foundation."

---

### CONCEPT SECTION

#### Opening — What Loops Are

**Paragraph 1:**
When a macro runs, it executes each line once, top to bottom. That works fine for
a handful of actions — but if you need to process 30 student records, you don't
want to write the same logic 30 times. Loops let you write the code once and repeat
it automatically for every record in your data.

**Paragraph 2:**
VBA has three loop structures: For Next, Do While, and Do Until. They all repeat
a block of code but differ in how they decide when to stop. For this course,
Do Until is the one you'll use most — it's built for processing rows of data
when you don't know in advance how many records there are.

**course-tip (concept):**
"The 'move to next record' line at the bottom of the loop is the one students
forget most often. If you forget it, the loop processes the same row forever.
Write that line before you write anything else inside the loop."

---

#### For Next Loop

**Paragraph 1:**
Use a For Next loop when you know exactly how many times the loop needs to run.
It uses a counter variable that starts at a value you set, runs the code inside,
increments by 1, then checks if it's reached the end value. If not, it loops again.
You can use any variable name for the counter — `i` is conventional — and you can
use variables for the start and end values instead of hardcoded numbers.

**.syntax-box:**
```
For [counter] = [start] To [end]
    ' your code here
Next [counter]
```

**Introduction sentence before code:**
"This example puts the numbers 1 through 4 into adjacent columns:"

```vba
Sub ForNextLoop()
    Sheets("Sheet2").Select
    Range("A1").Select
    Cells.ClearContents

    For i = 1 To 4
        ActiveCell.Value = i
        ActiveCell.Offset(0, i).Select
    Next i
End Sub
```

---

#### Do While Loop

**Paragraph 1:**
A Do While loop runs as long as a condition is true. VBA checks the condition
before each pass — if it's true, the code inside runs; if it's false, the loop
stops. If the condition is already false when the loop starts, the code inside
never runs at all. The condition must involve a variable that changes inside
the loop, otherwise the condition never becomes false and the loop runs forever.

**.syntax-box:**
```
Do While [condition is true]
    ' your code here
Loop
```

**Introduction sentence before code:**
"This example counts how many times it subtracts 1 from a number before
reaching 10:"

```vba
Sub ChkFirstWhile()
    counter = 0
    myNum = 20

    Do While myNum > 10
        myNum = myNum - 1
        counter = counter + 1
    Loop

    MsgBox "The loop made " & counter & " repetitions."
End Sub
```

---

#### Do Until Loop

**Paragraph 1:**
A Do Until loop runs until a condition becomes true — which means it keeps going
while the condition is false. For processing rows of data, the condition is always
a blank cell check: `Do Until ActiveCell = ""`. The loop continues as long as
there's data in the current cell and stops when it hits a blank row. This works
for 10 records or 10,000 without changing a single line of code.

**.syntax-box:**
```
Do Until [condition is true]
    ' your code here
Loop
```

**Introduction sentence before code:**
"The same counting example using Do Until — notice it reads as the inverse
of Do While:"

```vba
Sub ChkFirstUntil()
    counter = 0
    myNum = 20

    Do Until myNum = 10
        myNum = myNum - 1
        counter = counter + 1
    Loop

    MsgBox "The loop made " & counter & " repetitions."
End Sub
```

---

#### Processing All Records Pattern

**Paragraph 1:**
In this course, loops are almost always used to process rows of data — going through
a list of students or applicants and acting on each one. This pattern is the same
every time: navigate to the first record, loop until you hit a blank cell, do your
work, then move down one row at the very end of the loop body.

**.syntax-box:**
```
SELECT first record (e.g. Range("A2").Select)
DO UNTIL ActiveCell = ""
    ' process current row
    MOVE to next record (ActiveCell.Offset(1, 0).Select)
ENDLOOP
```

**Introduction sentence before code:**
"Here is what this pattern looks like in the actual Project Demo — the loop
that processes student applicants:"

```vba
Sheets("Applicant Information Group " & NewGroup & "").Select
Range("A2").Select

Do Until ActiveCell = ""

    If ActiveCell.Offset(0, 10) = Range("accept") Then

        UIN = ActiveCell
        GPR = ActiveCell.Offset(0, 1)

        Sheets("Student Information").Select
        ActiveCell = UIN
        ' ... populate remaining fields ...

        ActiveCell.Offset(1, 0).Range("PPAData[[#Headers],[UIN]]").Select
        Sheets("Applicant Information Group " & NewGroup & "").Select

    End If

    ' ALWAYS move to next record — whether accepted or not
    ActiveCell.Offset(1, 0).Range("ApplicantData[[#Headers],[UIN]]").Select

Loop
```

---

#### Cautions: Endless Loops

**Paragraph 1:**
An endless loop happens when the condition never becomes true. The three common
causes: the condition doesn't involve a variable, the variable never changes inside
the loop, or the "move to next record" line is inside an IF block instead of outside it.

**Bullet points:**
- The condition MUST involve a variable
- That variable MUST change inside the loop body
- The move line must be OUTSIDE any IF block — it must always execute
- If stuck: press **Ctrl+Break** to stop

---

#### AI Compare Panel

**Sanders panel label:** ✅ Sanders Approach
**Sanders panel h4:** Do Until — mirrors how you navigate in Excel

```vba
' Simple — easy to follow with F8
Sheets("Student Information").Select
Range("A2").Select

Do Until ActiveCell = ""
    ' process current row
    ActiveCell.Offset(1, 0).Select
Loop
```

**AI panel label:** ⚠️ Typical AI / Google Result
**AI panel h4:** For Each with object variables

```vba
' Uses concepts this course doesn't cover
Dim ws As Worksheet
Dim lastRow As Long
Dim cell As Range

Set ws = ThisWorkbook.Sheets("Student Information")
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For Each cell In ws.Range("A2:A" &amp; lastRow)
    If cell.Value &lt;&gt; "" Then
        ' process row
    End If
Next cell
```

**Explanation (.ai-compare-explanation):**
Both loops process every record — but the Sanders approach mirrors how you
navigate in Excel: start at A2, loop until blank, move down one row. You can
watch it step by step with F8 and see exactly where the cursor is at every line.

The AI version works, but it uses object variables, For Each, and End(xlUp) —
concepts this course doesn't cover. Using code you don't understand makes
debugging with F8 confusing and the exam significantly harder.

---

### QUICK CHECK SECTION

**Format:** Multiple choice. Lock on click. Reveal correct + explanation immediately.

**Question 1:**
You need to process every student record in a spreadsheet but you don't know
how many there are. Which loop should you use?
- A. For Next Loop
- B. Do Until Loop ← CORRECT
- C. Do While Loop
- D. It doesn't matter, they all work the same
**Explanation:** Do Until is the right choice when you don't know the record
count. `Do Until ActiveCell = ""` works for 10 records or 10,000 without
changing a single line.

**Question 2:**
In a Do Until loop, the code inside runs when the condition is _____.
- A. True
- B. False ← CORRECT
- C. Either true or false
- D. Only on the first pass
**Explanation:** Do Until runs while the condition is false, stops when
it becomes true. `Do Until ActiveCell = ""` keeps running while the cell
is NOT blank.

**Question 3:**
What happens if you forget `ActiveCell.Offset(1, 0).Select` at the bottom
of your Do Until loop?
- A. The loop skips every other row
- B. VBA throws an error and stops
- C. The macro runs forever on the same cell ← CORRECT
- D. The loop only runs once
**Explanation:** Without moving down, the loop checks the same cell
forever. The condition never becomes true. Press Ctrl+Break to stop it.

**Question 4:**
Which is the correct blank cell check for processing all records?
- A. `Do Until ActiveCell = 0`
- B. `Do Until ActiveCell = "done"`
- C. `Do Until ActiveCell = ""` ← CORRECT
- D. `Do Until ActiveCell.Row = 100`
**Explanation:** Checking for a blank cell works regardless of how many
records exist — no hardcoded row numbers needed.

**Question 5:**
In `For i = 1 To 4`, how many times does the loop run?
- A. 3 times
- B. 4 times ← CORRECT
- C. 5 times
- D. It depends on the data
**Explanation:** The loop runs for i = 1, 2, 3, and 4 — four times
inclusive.

**course-tip after quick check:**
"Question 2 catches people every semester. Do Until sounds like it runs
until something is true — but the code inside runs while the condition
is false. The loop asks 'am I done yet?' and keeps going while the answer is no."

---

### EASY WINS SECTION

#### Exercise 1 — Write Your First For Next Loop (STEPS FORMAT)
**Difficulty:** Guided

Write a macro that uses a For Next loop to put the numbers 1 through 5
into cells A1 through A5, one number per cell.

**Step 1 — Open the VBA Editor**
Press Alt+F11. Go to Insert → Module. You'll see a blank white area —
this is where you write code. Type this shell to start:
```vba
Option Explicit
Sub NumberCells()

End Sub
```
`Option Explicit` forces you to declare variables. Always include it.

**Step 2 — Navigate to your starting cell**
Add these two lines inside your Sub, before the loop:
```vba
Sheets("Sheet1").Select
Range("A1").Select
```
This selects Sheet1 and moves to A1 — your starting point.

**Step 3 — Write the For Next loop**
Add the loop inside your Sub:
```vba
For i = 1 To 5
    ActiveCell.Value = i
    ActiveCell.Offset(1, 0).Select
Next i
```
`ActiveCell.Value = i` puts the current number in the cell.
`ActiveCell.Offset(1, 0).Select` moves down one row.
`Next i` increases i by 1 and loops back.

**Step 4 — Run it**
Press F5 or the green Play button. Check Excel column A.

**Complete Code:**
```vba
Option Explicit
Sub NumberCells()
    Sheets("Sheet1").Select
    Range("A1").Select

    For i = 1 To 5
        ActiveCell.Value = i
        ActiveCell.Offset(1, 0).Select
    Next i
End Sub
```
**Expected result:** Cells A1–A5 contain 1, 2, 3, 4, 5.

---

#### Exercise 2 — Test the Do Until Condition (SIMPLE FORMAT)
**Difficulty:** Observation

Take the `ChkFirstUntil` example from above. Change `myNum = 20` to
`myNum = 10`. Run it. What happens? Why?

**Hint:** What does `Do Until myNum = 10` check before the loop starts?

**Solution:** The loop never runs — 0 repetitions. Because `myNum`
already equals 10 when the loop starts, the condition `myNum = 10`
is immediately true, so Do Until stops before executing a single line.
This is the key behavior: Do Until won't run at all if the condition
is already true at the start.

---

#### Exercise 3 — Count Non-Blank Cells with Do Until (STEPS FORMAT)
**Difficulty:** Guided

Write a macro that starts at A1, moves down the column using a Do Until
loop, and counts every non-blank cell. Display the count in a MsgBox.

**Step 1 — Set up your macro**
Open the VBA Editor (Alt+F11), add a module, create this shell:
```vba
Option Explicit
Sub CountCells()

End Sub
```

**Step 2 — Declare your variable**
Add this inside your Sub:
```vba
Dim CellCount As Integer
```
`Integer` is fine here — we won't exceed 32,000 rows.

**Step 3 — Navigate to your starting cell**
```vba
Sheets("Sheet1").Select
Range("A1").Select
```

**Step 4 — Write the Do Until loop**
```vba
Do Until ActiveCell = ""
    CellCount = CellCount + 1
    ActiveCell.Offset(1, 0).Select
Loop
```
`CellCount = CellCount + 1` adds 1 to the running total each pass.

**Step 5 — Display the result**
```vba
MsgBox "Non-blank cells found: " & CellCount
```

**Complete Code:**
```vba
Option Explicit
Sub CountCells()
    Dim CellCount As Integer

    Sheets("Sheet1").Select
    Range("A1").Select

    Do Until ActiveCell = ""
        CellCount = CellCount + 1
        ActiveCell.Offset(1, 0).Select
    Loop

    MsgBox "Non-blank cells found: " & CellCount
End Sub
```
**Expected result:** MsgBox shows the count of filled cells in column A.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table
**Instructions above toggle:**
"Copy this data and paste it into cell A1 of a new Excel worksheet.
Press Ctrl+V — Excel will split the columns automatically.
This is your dataset for the exercises below."

**30-record Aggie Advisors dataset** — full table from PRACTICE_PROJECT.md:
StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision
(all 30 rows as specified in PRACTICE_PROJECT.md)

#### Sample Data Exercise — Count Accepted Applicants
Using the Aggie Advisors data, write a macro that loops through all 30
applicant records and counts only those whose FinalDecision = "Accept".

**What your macro needs to do:**
- Start at cell A2 (row 1 is the header)
- Use `Do Until ActiveCell = ""` to loop through records
- Inside the loop: check if column H = "Accept"
  (FinalDecision is column H — Offset(0, 7) from column A)
- Add 1 to a counter for each Accept
- Move down one row at the bottom of the loop — every time, Accept or not
- After the loop: display "Accepted applicants: " & your counter

**Expected answer:** Accepted applicants: 20

**Hint:**
```vba
If ActiveCell.Offset(0, 7) = "Accept" Then
    AcceptCount = AcceptCount + 1
End If
' Move line goes AFTER End If, not inside it
ActiveCell.Offset(1, 0).Select
```

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-4"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Accepted and Denied GPR Averages
**No hints. No steps. Exam level.**

Using the same Aggie Advisors dataset, write a macro that:
1. Loops through all 30 records using Do Until with blank cell check
2. Separately accumulates total GPR for accepted and denied students
3. Counts accepted and denied students separately
4. Calculates the average GPR for each group after the loop
5. Displays both in a single MsgBox:
   "Accepted avg GPR: X.XX | Denied avg GPR: X.XX"

Must use Option Explicit. Must declare all variables with correct data types.

**Expected answers:**
- Accepted: 20 students, average GPR 3.7175
- Denied: 10 students, average GPR 3.0014

If your numbers don't match, use F8 and the Watch Window to step through
the first few records and check what your IF is actually comparing.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-4"`

---

## Remaining Modules — Practice Content

The practice content for each remaining module will be written here before
that module's task file is created. Never build a module page without the
full practice content being specified first.

### Module 1 — Macro Foundations

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Macro Foundations, you should have already watched
the Macro Foundations Video in Canvas and followed along with the Macro Demo file.
This practice will build upon that foundation."

---

## CONCEPT SECTION (id="concept")

### Opening — What a Macro Is

**Paragraph 1:**
A macro is a set of instructions that performs tasks in the order you specify.
Instead of clicking through the same sequence of actions every time, you record
or write those actions once and run them with a single click. Excel translates
each action into Visual Basic for Applications (VBA) code that you can view,
edit, and build on.

**Paragraph 2:**
You create macros in two ways: recording actions in Excel, or typing code
directly in the VBA Editor. Recording is how you start — it generates the
basic structure automatically. Typing is how you extend it with logic that
can't be recorded, like IF statements, variables, and loops. Most macros
in this course use both.

**course-tip (concept):**
"The recorder is your starting point, not your limitation. Record the actions,
then open the code and understand what it generated. That's the workflow
everything else builds on."

---

#### The VBA Editor

**Paragraph 1:**
The VBA Editor is where your macro code lives. Press **Alt+F11** to open it.
You'll see a Project panel on the left showing your workbook and its modules —
modules are where macros are stored. Double-click a module to see its code.
The code window is where you read, write, and edit macros.

**.syntax-box:**
```
Alt+F11          → Open/close the VBA Editor
F5               → Run the current macro
F8               → Step through code one line at a time
Alt+F11 again    → Switch back to Excel
```

**Introduction sentence before code:**
"Every macro follows this basic structure:"

```vba
Sub MacroName()
    ' Your recorded or typed code goes here
End Sub
```

---

#### Recorded vs Typed Code

**Paragraph 1:**
Actions you perform in Excel — selecting cells, formatting, copying, navigating —
can be recorded. Programming concepts that extend what Excel can do — IF statements,
variables, loops, and InputBoxes — must be typed. Understanding which is which
helps you know when to hit Record and when to open the editor and write.

**.syntax-box:**
```
Recorded (actions in Excel):
  Range("A1").Select
  Selection.Font.Bold = True
  ActiveSheet.Name = "Report"

Typed (programming concepts):
  If ActiveCell = "" Then ...
  Dim Counter As Integer
  Do Until ActiveCell = ""
  InputBox("Enter a value")
```

---

#### Saving Your File

**Paragraph 1:**
A regular Excel file (.xlsx) cannot save macros. When you close the file,
any macros you recorded are gone. You must save as a **Macro-Enabled Workbook
(.xlsm)** to keep your code. Excel will warn you if you try to save a file
with macros as .xlsx — always choose to keep the macro-enabled format.

**.syntax-box:**
```
File → Save As → Excel Macro-Enabled Workbook (.xlsm)
```

**One more thing:** You cannot undo (Ctrl+Z) changes made by a macro.
Before running a new or untested macro, save a backup of your file.

---

### QUICK CHECK SECTION (id="quick-check")

**Format:** Multiple choice. Lock on click. Immediate feedback.

**Question 1:**
What keyboard shortcut opens the VBA Editor?
- A. Ctrl+V
- B. Alt+F8
- C. Alt+F11 ← CORRECT
- D. Ctrl+F11
**Explanation:** Alt+F11 opens and closes the VBA Editor. Alt+F8 opens
the macro list where you can run existing macros.

**Question 2:**
Which of the following must be TYPED and cannot be recorded?
- A. Selecting a cell
- B. Bolding a row
- C. Renaming a worksheet
- D. An IF statement ← CORRECT
**Explanation:** IF statements, variables, loops, and InputBoxes are
programming concepts — they must be typed. Actions you perform in Excel
(selecting, formatting, navigating) can be recorded.

**Question 3:**
You record a macro in an .xlsx file and close Excel. What happens to the macro?
- A. It is saved with the file
- B. It is lost — .xlsx cannot save macros ← CORRECT
- C. It moves to the Personal Workbook automatically
- D. Excel converts the file to .xlsm automatically
**Explanation:** .xlsx files cannot store macros. Always save as .xlsm
(Excel Macro-Enabled Workbook) before closing if you want to keep your code.

**Question 4:**
You run a macro that accidentally deletes some data. You press Ctrl+Z. What happens?
- A. The data is restored
- B. Nothing — macro changes cannot be undone ← CORRECT
- C. The macro runs again in reverse
- D. Excel asks if you want to restore the data
**Explanation:** Ctrl+Z does not work on macro changes. Always keep a
backup before running a new or untested macro.

**Question 5:**
Where are macros stored inside a workbook?
- A. In a hidden worksheet
- B. In the cell comments
- C. In modules inside the VBA Editor ← CORRECT
- D. In the file properties
**Explanation:** Macros live in modules, which you can see in the Project
panel on the left side of the VBA Editor. Double-click a module to view its code.

**course-tip after quick check:**
"Questions 3 and 4 are the ones students learn the hard way — losing a macro
because they saved as .xlsx, or losing data because Ctrl+Z didn't work.
Both are easy to avoid once you know."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Record Your First Macro (STEPS FORMAT)
**Difficulty:** Guided

Record a macro that types your name into cell A1 and bolds it.
This is the simplest possible macro — the goal is just to see
the whole process from recording to running to viewing the code.

**Step 1 — Set up a new workbook**
Open Excel. Create a new blank workbook. Save it immediately as
an .xlsm file: File → Save As → Excel Macro-Enabled Workbook.
Name it whatever you like. If you skip this step and save as .xlsx
later, your macro will be lost.

**Step 2 — Start recording**
Go to the **Developer** tab → **Record Macro**.
(If you don't see the Developer tab: File → Options → Customize Ribbon
→ check Developer → OK)

In the dialog box:
- Macro name: TypeMyName (no spaces allowed in macro names)
- Store macro in: This Workbook
- Click OK — recording has started

**Step 3 — Perform the actions**
Click cell A1. Type your name. Press Enter.
Click cell A1 again. Press Ctrl+B to bold it.
That's it — keep it simple.

**Step 4 — Stop recording**
Go to Developer → **Stop Recording**.
Your macro is now saved.

**Step 5 — View the code**
Press Alt+F11 to open the VBA Editor.
In the Project panel, expand your workbook → Modules → Module1.
Double-click it. You should see the code Excel generated.

**Step 6 — Run it on a new cell**
Close the VBA Editor (Alt+F11). Click cell B1. Delete cell A1's content.
Go to Developer → Macros → TypeMyName → Run.
Watch what happens.

**Complete Code (what you should see — yours may vary slightly):**
```vba
Sub TypeMyName()
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Your Name"
    Range("A1").Select
    Selection.Font.Bold = True
End Sub
```
**What to notice:** Excel always records a Select first, then performs
the action on the Selection. Later modules will show you how to clean
this up — for now, just observe what the recorder generated.

---

#### Exercise 2 — Break It and Fix It (SIMPLE FORMAT)
**Difficulty:** Observation

Open the macro you just recorded. Change the cell reference from
`Range("A1")` to `Range("C3")` in both places. Run it. What happens?

**Hint:** You're editing the hardcoded cell address. This is what
"absolute reference" means in recorded code — it always goes to
the same cell no matter where you are when you run it.

**Solution:** The macro now types your name in C3 and bolds it,
regardless of which cell was selected before you ran it. The reference
is hardcoded — it doesn't adapt to your position. This is the difference
between absolute and relative references, which gets its own module later.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table
No data table needed for this module — the practice problem uses
the student's own Excel file from the Easy Wins exercise.

#### Practice Problem — Record a Useful Macro
**No data table toggle needed. Use .sample-data-exercise directly.**

Think of something repetitive you do in Excel — even something small.
Record a macro that does it. Some ideas if you're not sure:

- Navigate to a specific sheet and select cell A1
- Bold the first row and autofit all columns
- Add today's date to a specific cell
- Clear the contents of a range

**What your macro needs to do:**
- Be recorded using Developer → Record Macro
- Be saved in This Workbook
- Actually run and do what you intended when you press Play

**After recording:**
- Open the VBA Editor and look at the code
- Find at least one line you can explain in plain English
- Find at least one line you're not sure about — note it down
  for when you get to later modules

There is no single correct answer here. The goal is to get comfortable
with the record → view → run cycle before adding any complexity.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-1"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Record, View, and Describe
**No hints. Exam level.**

Record a macro that does ALL of the following in one recording:
1. Navigates to Sheet2 (create it if it doesn't exist)
2. Types "Sales Report" into cell A1
3. Bolds cell A1
4. Types today's date into cell A2
5. Returns to Sheet1

After recording, open the VBA Editor and answer these questions
by looking at the generated code:

- How many times does the word "Select" appear?
- What line puts the text "Sales Report" into the cell?
- What does `ActiveCell.FormulaR1C1` mean based on what you see?
- Is the cell reference in your code absolute or relative?
  How can you tell?

Write your answers in a comment block at the top of the macro
using the VBA comment format (apostrophe at the start of each line).

**Expected result:** A working macro that performs all 5 steps,
with a comment block at the top showing your answers.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-1"`

### Module 2 — Adding Programming Concepts

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Adding Programming Concepts, you should have
already watched the Adding Programming Concepts Video in Canvas and followed
along with the Macro Demo file. This practice will build upon that foundation."

---

## CONCEPT SECTION (id="concept")

### Opening — Beyond Recording

**Paragraph 1:**
Recording captures actions — clicking, typing, navigating. But macros become
truly useful when they can make decisions, ask the user for input, and handle
different situations automatically. These capabilities can't be recorded. They
have to be typed. This module covers the three essentials: Option Explicit,
InputBox, and the guard check pattern.

**Paragraph 2:**
The workflow from here forward is always the same: record the parts you can,
open the code, and add the typed logic around it. You're not writing everything
from scratch — you're extending what the recorder gives you.

**course-tip (concept):**
"The record → open → add workflow is exactly how the project is built.
Practice this rhythm now and the project will feel familiar rather than overwhelming."

---

#### Option Explicit

**Paragraph 1:**
Option Explicit forces you to declare every variable before using it. Without it,
VBA silently creates a new variable any time it sees an unrecognized word —
including when you misspell a variable name. That misspelled variable gets a blank
value and your macro produces wrong results with no error message. Option Explicit
must be the very first line in the module, above the first Sub.

**.syntax-box:**
```
Option Explicit     ← first line in the module, above everything

Sub MacroName()
    Dim VariableName As DataType
    ' ...
End Sub
```

**Introduction sentence before code:**
"This is what every module in this course should start with:"

```vba
Option Explicit
' ACCT 628 - Sanders

Sub AddNewStudents()

    Dim NewGroup       As Integer
    Dim NumberAccepted As Integer

End Sub
```

---

#### InputBox

**Paragraph 1:**
An InputBox pauses the macro and displays a dialog asking the user to type
something. Whatever they type gets stored in a variable. This is how macros
become adaptable — instead of hardcoding a group number or a date, you ask
for it at runtime. InputBox is always typed, never recorded.

**.syntax-box:**
```
VariableName = InputBox("Message to display to the user")
```

**Introduction sentence before code:**
"This example asks for a group number and stores it in a variable:"

```vba
Dim NewGroup As Integer

NewGroup = InputBox("Enter the new group number")
```

---

#### Speed Settings

**Paragraph 1:**
Two settings slow macros down significantly when left on: automatic calculation
and screen updating. Turning them off at the start of a macro and back on at the
end is standard practice for any macro that processes a large number of rows.
Always add them as a pair — if you turn calculation off and forget to turn it
back on, formulas in the workbook will stop updating until you fix it manually.

**.syntax-box:**
```
' At the start of your macro
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

' At the end of your macro
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

---

#### The Guard Check Pattern

**Paragraph 1:**
Before a macro processes data, it should verify that data actually exists.
If the table is empty and the macro tries to loop through it, you'll get
unexpected results. The guard check is a simple IF statement at the top of
the macro: if the first data cell is blank, show a message and stop. If it
has data, continue. This one pattern prevents a whole category of errors.

**.syntax-box:**
```
IF first data cell is empty THEN
    DISPLAY message and stop
ELSE
    run the rest of the macro
END IF
```

**Introduction sentence before code:**
"Here is what the guard check looks like in VBA:"

```vba
Sheets("Applicant Information").Select
Range("A2").Select

If ActiveCell = "" Then
    MsgBox "No applicants found. Please enter data first."
    Exit Sub
Else
    ' rest of macro goes here
End If
```

---

### QUICK CHECK SECTION (id="quick-check")

**Format:** Multiple choice. Lock on click. Immediate feedback.

**Question 1:**
What does Option Explicit do?
- A. Makes your macro run faster
- B. Forces you to declare all variables before using them ← CORRECT
- C. Turns off screen updating
- D. Opens the VBA Editor automatically
**Explanation:** Option Explicit catches misspelled variable names by
requiring every variable to be declared with Dim. Without it, a typo
silently creates a blank variable and produces wrong results.

**Question 2:**
Where must Option Explicit be placed?
- A. Inside the Sub, before the Dim statements
- B. After the last End Sub in the module
- C. At the very top of the module, above the first Sub ← CORRECT
- D. It can go anywhere
**Explanation:** Option Explicit must be the very first line in the module.
If it's inside a Sub it will cause an error.

**Question 3:**
What does this line do?
`NewGroup = InputBox("Enter the new group number")`
- A. Displays a message and continues
- B. Pauses the macro and stores whatever the user types into NewGroup ← CORRECT
- C. Populates a cell with the text "Enter the new group number"
- D. Declares a variable named NewGroup
**Explanation:** InputBox pauses the macro, shows the message to the user,
and stores their input in the variable on the left side of the equals sign.

**Question 4:**
You turn off Application.Calculation at the start of your macro but forget
to turn it back on. What happens?
- A. Nothing — it resets automatically when the macro ends
- B. The macro throws an error
- C. Formulas in the workbook stop recalculating until you fix it manually ← CORRECT
- D. Excel saves the file automatically to compensate
**Explanation:** Calculation stays off until you explicitly set it back to
xlCalculationAutomatic. Formulas will show a strikethrough. Fix it in Excel
via Formulas → Calculation Options → Automatic.

**Question 5:**
In the guard check pattern, what is `Exit Sub` used for?
- A. It closes the VBA Editor
- B. It stops the macro from running any further ← CORRECT
- C. It exits the current IF block only
- D. It saves the workbook
**Explanation:** Exit Sub stops the macro immediately and exits the Sub.
In the guard check, it's used to stop the macro gracefully when the data
table is empty instead of letting it run and produce wrong results.

**course-tip after quick check:**
"The guard check and the Calculation on/off pair both show up on the exam
in the form of 'what's wrong with this code' questions. Know what happens
when each one is missing."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Add Option Explicit and See What Happens (STEPS FORMAT)
**Difficulty:** Guided

Add Option Explicit to an existing macro and intentionally misspell a
variable name to see exactly what error you get and why it helps.

**Step 1 — Open a macro from the previous module**
Open the workbook from Module 1 (or any workbook with a macro).
Press Alt+F11 to open the VBA Editor.

**Step 2 — Add Option Explicit**
At the very top of the module — above the Sub line — type:
```vba
Option Explicit
```
Nothing should break yet.

**Step 3 — Add a variable and misspell it**
Inside your existing Sub, add these two lines:
```vba
Dim MyValue As Integer
MyVlaue = 10
```
Notice the deliberate misspelling: MyVlaue instead of MyValue.

**Step 4 — Run the macro**
Press F5. VBA will stop and highlight the misspelled line with an error:
"Variable not defined."

Without Option Explicit, VBA would have silently created a new blank
variable called MyVlaue and MyValue would have stayed 0. No error,
just wrong results. Option Explicit caught it immediately.

Fix: Correct the spelling to MyValue = 10. Run again — no error.

---

#### Exercise 2 — Add an InputBox (SIMPLE FORMAT)
**Difficulty:** Guided

Add an InputBox to the macro from Exercise 1. Ask the user for a number,
store it in a variable, and display it back in a MsgBox.

```vba
Option Explicit
Sub InputBoxDemo()
    Dim UserNumber As Integer
    UserNumber = InputBox("Enter a number:")
    MsgBox "You entered: " & UserNumber
End Sub
```

**Hint:** Run it. Type a number. See what the MsgBox shows.
Then run it again and type a word instead of a number. What happens?
(Integer can't store text — VBA will show a type mismatch error.
This is why data types matter.)

**Solution:** The macro works correctly when you enter a number.
When you type text into an Integer variable, VBA throws a Type Mismatch
error. This demonstrates why choosing the right data type is important —
and why String is safer when you're not sure what the user will type.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table (Aggie Advisors — 30 records)
**Instructions above toggle:**
"Copy this data and paste it into cell A1 of a new Excel worksheet.
Press Ctrl+V — Excel will split the columns automatically."

Full 30-record dataset from PRACTICE_PROJECT.md:
StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision

#### Practice Problem — Add a Guard Check
Using the Aggie Advisors data, write a macro that checks whether the
Applicant Information sheet has any records before doing anything else.

**What your macro needs to do:**
- Use Option Explicit
- Navigate to the Applicant Information sheet and select cell A2
- Check if A2 is empty using an IF statement
- If empty: display "No applicants found." and stop with Exit Sub
- If not empty: prompt the user for a group number using InputBox,
  store it in an Integer variable, and display:
  "Group [X] is ready to process" where X is the number entered

**Test it both ways:**
1. Run it with data in A2 — you should get the InputBox and then the message
2. Clear cell A2, run it again — you should get "No applicants found."

**Expected results:**
- With data: InputBox appears, then "Group [X] is ready to process"
- Without data: "No applicants found." and macro stops

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-2"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Complete Macro Setup
**No hints. Exam level.**

Write a macro from scratch that includes all of the following:
1. Option Explicit at the top of the module
2. Speed settings turned off at the start and back on at the end
3. A guard check that stops with an appropriate message if A2 on
   the Applicant Information sheet is empty
4. An InputBox that asks for a group number (Integer)
5. A MsgBox at the end that displays:
   "Setup complete. Ready to process Group [X]"
   where X is the number from the InputBox

The macro must handle both cases cleanly — empty table stops gracefully,
data present continues to the completion message.

**Expected behavior:**
- Empty table: "No applicants found." — macro stops, speed settings restored
- Data present: InputBox → "Setup complete. Ready to process Group [X]"

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-2"`

### Module 3 — Variables

### Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Variables, you should have already watched the
Variables Video in Canvas and followed along with the Macro Demo file.
This practice will build upon that foundation."

---

### CONCEPT SECTION

#### Opening — What Variables Are

**Paragraph 1:**
A variable is a named storage location in your computer's memory. Instead of
hardcoding a value like a group number or GPR directly into your code, you store
it in a variable and reference the variable name instead. This makes your macro
adaptable — change the value once and it updates everywhere the variable is used.

**Paragraph 2:**
Working with variables always follows three steps. Professor Sanders calls them
Define, Populate, and Use — and that order is not optional. You must define a
variable before you can populate it, and populate it before you can use it.
If a value is wrong or blank, the first thing to check is whether these three
steps are happening in the right order.

**course-tip (concept):**
"Variables show up in every exam question in some form. Getting the three steps
and data types locked in now means one less thing to think about on exam day."

---

#### The Three Steps

**Paragraph 1:**
Defining a variable creates a named slot in memory and tells VBA what kind of
data it will hold. Populating it places a value in that slot. Using it means
referencing the variable name anywhere you need that value — in a cell assignment,
a MsgBox, or a loop condition.

**.syntax-box:**
```
' Step 1 — DEFINE
Dim VariableName As DataType

' Step 2 — POPULATE
VariableName = something

' Step 3 — USE
ActiveCell = VariableName
```

**Introduction sentence before code:**
"Here is the three-step pattern in a real macro — prompting for a group number
and writing it to a cell:"

```vba
Dim NewGroup As Integer

NewGroup = InputBox("Enter the new group number")

Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveCell = NewGroup
```

---

#### Data Types

**Paragraph 1:**
The data type tells VBA how to store the value in memory. Using the wrong type
causes silent errors — an Integer silently truncates a decimal, a String won't
compare correctly to a number. For this course you need five types: String for
text, Integer for small whole numbers, Long for large whole numbers like student
IDs, Double for numbers with decimal places like GPR, and Date for dates.

**.syntax-box:**
```
Dim LastName    As String    ' text
Dim GroupNumber As Integer   ' whole number up to ~32,000
Dim StudentID   As Long      ' whole number any size
Dim GPR         As Double    ' number with decimals
Dim StartDate   As Date      ' date value
```

**Introduction sentence before code:**
"These are the actual variable declarations from the Project Demo:"

```vba
Dim NewGroup       As Integer   ' group number — small whole number
Dim NumberAccepted As Integer   ' counter — small whole number
Dim UIN            As Long      ' 9-digit student ID — needs Long
Dim GPR            As Double    ' GPA with decimals
Dim Gender         As String    ' text
Dim Birthdate      As Date      ' date
```

---

#### Option Explicit

**Paragraph 1:**
By default, VBA lets you use any word as a variable without declaring it.
This causes three problems: VBA may choose the wrong data type, a misspelled
variable name silently creates a new blank variable instead of an error, and
your macro runs slower. Adding `Option Explicit` at the very top of the module
— above the first Sub — forces you to declare every variable. If you use an
undeclared name, VBA stops and tells you exactly where the problem is.

**.syntax-box:**
```
Option Explicit          ' ← must be the very first line in the module

Sub YourMacroName()
    Dim VariableName As DataType
End Sub
```

**Introduction sentence before code:**
"This is what the top of every well-written module looks like:"

```vba
Option Explicit
' ACCT 628 - Sanders

Sub Option1_Modified()

    Dim NewGroup       As Integer
    Dim NumberAccepted As Integer
    Dim UIN            As Long
    Dim GPR            As Double
    Dim Gender         As String
    Dim Birthdate      As Date
```

---

#### Variables vs Named Ranges

**Paragraph 1:**
Variables and Named Ranges look similar but behave very differently. A variable
exists only while the macro is running — when the macro stops, it's gone.
A Named Range exists on the worksheet permanently and can be used in both
Excel formulas and VBA. In VBA, Named Ranges always use quotes:
`Range("GroupData")`. Variables never use quotes.

**.syntax-box:**
```
' Variable — no quotes, exists only while macro runs
NewGroup = InputBox("Enter group number")

' Named Range — quotes required, exists on worksheet
Range("LegalDrinkingAge") = LegalDrinkingAge
```

**Introduction sentence before code:**
"This example from the Macro Handout shows both in the same macro:"

```vba
Dim LegalDrinkingAge As Integer

LegalDrinkingAge = InputBox("Enter the current legal drinking age")

' Named Range has quotes — Variable does not
Range("LegalDrinkingAge") = LegalDrinkingAge
```

---

### QUICK CHECK SECTION

**Format:** Multiple choice. Lock on click. Immediate feedback.

**Question 1:**
What are the three steps for working with variables? Use Professor Sanders' exact terms.
- A. Create, Assign, Reference
- B. Define, Populate, Use ← CORRECT
- C. Declare, Set, Call
- D. Dim, Equal, Run
**Explanation:** Define (Dim), Populate (VariableName = something), Use
(reference it in code). This order is not optional.

**Question 2:**
A student's UIN is a 9-digit number like 123456789. Which data type should you use?
- A. Integer
- B. String
- C. Double
- D. Long ← CORRECT
**Explanation:** Integer only holds numbers up to ~32,000. A 9-digit UIN needs
Long, which handles numbers up to ~2 billion.

**Question 3:**
You declare `Dim GPR As Integer` and assign it 3.756. What value does GPR contain?
- A. 3.756
- B. 3
- C. 4 ← CORRECT
- D. An error message
**Explanation:** Integer silently rounds to the nearest whole number — 3.756
becomes 4. No error, no warning. Use Double for any value with decimal places.

**Question 4:**
Where must `Option Explicit` be placed?
- A. Inside the Sub, before the Dim statements
- B. At the very top of the module, above the first Sub ← CORRECT
- C. After the last End Sub
- D. It can go anywhere
**Explanation:** Option Explicit must be the first line in the module, above
any Sub. If it's inside a Sub it will cause an error.

**Question 5:**
In VBA, how do you reference a Named Range called "GroupData"?
- A. GroupData
- B. Variable("GroupData")
- C. Range("GroupData") ← CORRECT
- D. Named("GroupData")
**Explanation:** Named Ranges always use Range("name") with quotes in VBA.
Variables never use quotes. This distinction shows up on the exam.

**course-tip after quick check:**
"Question 3 is the classic trap. Integer looks right until you assign a decimal
and it silently rounds. Always ask: could this value have a decimal? If yes, use Double."

---

### EASY WINS SECTION

#### Exercise 1 — Your First Variable (STEPS FORMAT)
**Difficulty:** Guided

Write a macro that asks for a name using InputBox, stores it in a String
variable, and displays "Hello, [name]!" in a MsgBox.

**Step 1 — Set up your macro**
Open the VBA Editor (Alt+F11). Insert a module (Insert → Module). Type:
```vba
Option Explicit
Sub SayHello()

End Sub
```

**Step 2 — Define your variable**
Inside the Sub, add:
```vba
Dim UserName As String
```
String is correct because a name is text.

**Step 3 — Populate the variable**
```vba
UserName = InputBox("What is your name?")
```
This pauses the macro, waits for input, and stores the result in UserName.

**Step 4 — Use the variable**
```vba
MsgBox "Hello, " & UserName & "!"
```
The `&` operator joins text and variable values into one string.

**Complete Code:**
```vba
Option Explicit
Sub SayHello()
    Dim UserName As String
    UserName = InputBox("What is your name?")
    MsgBox "Hello, " & UserName & "!"
End Sub
```
**Expected result:** MsgBox shows "Hello, [whatever you typed]!"

---

#### Exercise 2 — Data Type Experiment (SIMPLE FORMAT)
**Difficulty:** Observation

Declare two variables — one Integer, one Double. Assign 3.7 to both.
Display both in a MsgBox. What's different and why?

```vba
Dim WholeNumber   As Integer
Dim DecimalNumber As Double

WholeNumber   = 3.7
DecimalNumber = 3.7

MsgBox "Integer: " & WholeNumber & " | Double: " & DecimalNumber
```

**Hint:** Run it and look at what Integer does to 3.7.

**Solution:** Integer shows 4 (rounds to nearest whole number). Double shows 3.7.
This is why you always use Double for GPR, dollar amounts, or any decimal value.

---

#### Exercise 3 — Three Steps in Order (STEPS FORMAT)
**Difficulty:** Guided

Write a macro that reads a student's ID, last name, and GPR from cells
A2, B2, and D2, stores them in variables, then displays all three in a MsgBox.

**Step 1 — Define three variables**
```vba
Option Explicit
Sub ReadStudent()
    Dim StudentID As Long
    Dim LastName  As String
    Dim GPR       As Double
```

**Step 2 — Navigate and populate**
```vba
    Sheets("Sheet1").Select
    Range("A2").Select

    StudentID = ActiveCell
    LastName  = ActiveCell.Offset(0, 1)
    GPR       = ActiveCell.Offset(0, 3)
```
Offset(0,1) = one column right. Offset(0,3) = three columns right.

**Step 3 — Use the variables**
```vba
    MsgBox "ID: " & StudentID & " | Name: " & LastName & " | GPR: " & GPR
End Sub
```

**Complete Code:**
```vba
Option Explicit
Sub ReadStudent()
    Dim StudentID As Long
    Dim LastName  As String
    Dim GPR       As Double

    Sheets("Sheet1").Select
    Range("A2").Select

    StudentID = ActiveCell
    LastName  = ActiveCell.Offset(0, 1)
    GPR       = ActiveCell.Offset(0, 3)

    MsgBox "ID: " & StudentID & " | Name: " & LastName & " | GPR: " & GPR
End Sub
```
**Expected result:** MsgBox shows values from row 2 of your data.

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table (Aggie Advisors — 30 records)
**Instructions above toggle:**
"Copy this data and paste it into cell A1 of a new Excel worksheet.
Press Ctrl+V — Excel will split the columns automatically."

Full 30-record dataset from PRACTICE_PROJECT.md:
StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision

#### Practice Problem — Read the First Applicant
Using the Aggie Advisors data, write a macro that reads the first applicant
record and displays their information in a MsgBox.

**What your macro needs to do:**
- Declare four variables: StudentID (Long), LastName (String),
  FirstName (String), GPR (Double)
- Navigate to cell A2
- Populate each variable using Offset:
  - StudentID = ActiveCell (column A)
  - LastName = Offset(0, 1) (column B)
  - FirstName = Offset(0, 2) (column C)
  - GPR = Offset(0, 3) (column D)
- Display: "ID: [X] | Name: [Last], [First] | GPR: [X.XXX]"

**Expected result for row 2:**
"ID: 724816395 | Name: Anderson, Emma | GPR: 3.842"

**Hint:**
```vba
MsgBox "ID: " & StudentID & " | Name: " & LastName & ", " & FirstName & " | GPR: " & GPR
```

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-3"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Read and Classify All Applicants
**No hints. No steps. Exam level.**

Using the Aggie Advisors dataset, write a macro that:
1. Uses Option Explicit with all variables declared and correct data types
2. Loops through all 30 applicant records using Do Until with blank cell check
3. For each record, reads StudentID, LastName, and FinalDecision into variables
4. Counts how many students have FinalDecision = "Accept"
5. After the loop, displays: "Accepted: [X] of [Total] applicants"
   where Total comes from your loop counter, not a hardcoded 30

**Expected answer:** "Accepted: 20 of 30 applicants"

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-3"`

### Module 5 — Calculations and Dates
[To be written before calculations.html is built]

### Module 6 — Relative vs Absolute References
[To be written before references.html is built]

### Module 7 — Filters & Shortcut Keys
[To be written before filters.html is built]

### Module 8 — F8 Debugging Practice

## Canvas prerequisite (.box-reminder):
"REMINDER: To get the most out of this module, you should have already watched
the Macro Demo videos in Canvas and followed along with the Macro Demo file.
This module is different from the others — the entire point is to practice
a technique, not read about it. Open Excel and VBA before you start."

---

## PAGE STRUCTURE NOTE

This module uses a modified section structure compared to other modules.
The anchor nav has these five pills:
- Concept (#concept)
- Setup (#setup)
- Walkthrough (#walkthrough)
- On Your Own (#on-your-own)
- Aggie Advisors (#challenge)

The .page-subtitle is "Module 8 of 9".

---

## CONCEPT SECTION (id="concept")

### Opening — What F8 Actually Does

**Paragraph 1:**
When a macro doesn't work, most students read the code looking for the problem.
That rarely works. F8 does something completely different — it runs your macro
one line at a time and shows you what Excel actually does at each step. You watch
the cursor move, watch variables change value, and watch conditions evaluate in
real time. The bug reveals itself because you can see exactly where the code
stops doing what you expected.

**Paragraph 2:**
Professor Sanders said it best: lightbulbs go off when students start doing this.
It changes how you think about VBA. A macro isn't a wall of text — it's a sequence
of actions you can step through and observe. Once you've used F8 to find a bug,
reading code passively feels like trying to find a spelling mistake by staring
at a paragraph instead of reading it word by word.

**course-tip (concept):**
"If you only learn one thing from this site, make it F8. Every exam question
that asks 'what does this macro do' becomes straightforward once you've
practiced stepping through code and watching what actually happens."

---

## SETUP SECTION (id="setup")

### h2: Setting Up Your Screen

**Paragraph 1:**
The key to F8 debugging is having both windows visible at the same time — the
VBA Editor on one side, Excel on the other. When you step through code, you need
to see where the cursor moves in Excel and watch variable values change in the
Watch Window simultaneously. If you can only see one window at a time, you're
missing half the information.

**How to set up (numbered list — not a steps component, just an ol):**
1. Open Excel and press **Alt+F11** to open the VBA Editor
2. In the VBA Editor, go to **Window → Tile Vertically** (or drag the VBA window
   to one side of your screen and Excel to the other)
3. Open the Watch Window: **View → Watch Window** in the VBA Editor
4. You should now see: Excel on one side, VBA Editor + Watch Window on the other

**How to add a Watch (numbered list):**
1. In your code, highlight the variable name you want to watch
2. Drag it down to the Watch Window — or right-click → Add Watch
3. The Watch Window shows the variable's current value and data type
4. Add both sides of an IF comparison when debugging a condition — seeing
   what each side actually contains is how you catch type mismatches

**How to set a Breakpoint (numbered list):**
1. Click in the grey margin bar to the left of the line where you want to pause
2. A red dot appears — this is your breakpoint
3. Press **F5** (Play) to run the macro — it will stop at the breakpoint
4. Then use **F8** to step one line at a time from that point
5. Press **F9** to toggle a breakpoint on/off, or click the red dot to remove it

**course-tip (setup):**
"Set your breakpoint at the start of the loop, not the start of the macro.
Running full-speed to the loop and then stepping through is much faster than
F8-ing through every line of setup code to get there."

---

## WALKTHROUGH SECTION (id="walkthrough")

### h2: Guided Walkthrough — Find the Bug

**Intro paragraph:**
The macro below runs without an error message but gives the wrong answer.
Your job is to find the bug using F8 and the Watch Window. Follow each step
exactly — this is the technique you'll use on the exam.

### The Broken Macro

**Instructions:** Copy this macro into a new module in Excel. This is a modified
version of the CumulativeTotal macro from the Macro Handout. It should add up
all the values in column F and display the total — but it gives the wrong answer.
Don't try to spot the bug by reading. Follow the steps.

Assume your data is in a worksheet named "Totals" with numeric values in column F
starting at row 2. Use any numeric data you have, or create 5 rows with values
like 100, 200, 300, 400, 500 in column F.

```vba
Option Explicit
Sub CumulativeTotal_Broken()

    Dim CumulativeTotal As Double

    Sheets("Totals").Select
    Range("A2").Select

    Do Until ActiveCell = ""
        CumulativeTotal = ActiveCell.Offset(0, 5)
        ActiveCell.Offset(1, 0).Select
    Loop

    MsgBox "The cumulative total is " & CumulativeTotal

End Sub
```

### Steps (STEPS FORMAT — .exercise-steps)

**Step 1 — Run it first and note the wrong answer**
Press F5 to run the macro normally. Write down the number in the MsgBox.
If your column F has 100, 200, 300, 400, 500 — the correct total is 1500.
The macro will show 500 instead. Now you know what you're looking for.

**Step 2 — Add a Watch for CumulativeTotal**
In the VBA Editor, highlight the word `CumulativeTotal` anywhere in the code.
Drag it to the Watch Window. You should see it appear with Value = 0 and
Type = Double. Now you can watch it change as the macro runs.

**Step 3 — Set a breakpoint at the start of the loop**
Click the grey margin bar to the left of the `Do Until ActiveCell = ""` line.
A red dot appears. This is where the macro will pause when you press Play.

**Step 4 — Run and step through the first two iterations**
Press F5. The macro pauses at your breakpoint. Now press F8 three times slowly:
- First F8: the Do Until condition evaluates — watch the cursor jump to Excel
- Second F8: the `CumulativeTotal = ActiveCell.Offset(0, 5)` line executes —
  look at the Watch Window. What is CumulativeTotal now?
- Third F8: `ActiveCell.Offset(1, 0).Select` executes — the cursor moves down

Press F8 two more times to complete the second loop iteration.
What is CumulativeTotal after the second iteration? It should be 300 if
accumulating correctly (100 + 200). Is it?

**Step 5 — Identify the bug and fix it**
You should have noticed that CumulativeTotal after iteration 1 = 100, and
after iteration 2 = 200 — not 300. It's replacing the total each time instead
of adding to it.

The bug is on this line:
```vba
CumulativeTotal = ActiveCell.Offset(0, 5)
```

It should be:
```vba
CumulativeTotal = CumulativeTotal + ActiveCell.Offset(0, 5)
```

Make that one change. Clear the breakpoint (click the red dot). Press F5.
The MsgBox should now show 1500.

**Complete Fixed Code:**
```vba
Option Explicit
Sub CumulativeTotal_Fixed()

    Dim CumulativeTotal As Double

    Sheets("Totals").Select
    Range("A2").Select

    CumulativeTotal = 0

    Do Until ActiveCell = ""
        CumulativeTotal = CumulativeTotal + ActiveCell.Offset(0, 5)
        ActiveCell.Offset(1, 0).Select
    Loop

    MsgBox "The cumulative total is " & CumulativeTotal

End Sub
```

**What you just practiced:**
- Setting a breakpoint to pause at a specific line
- Using the Watch Window to monitor a variable's value
- Stepping with F8 to see what each line actually does
- Identifying a logic error by watching the variable NOT change the way you expected

---

## ON YOUR OWN SECTION (id="on-your-own")

### h2: On Your Own — Debug Without Guidance

**Intro paragraph:**
Same technique, no step-by-step this time. This macro should find all students
with a GPR above 3.5 and count them — but it runs and always shows 0.
Use F8 and the Watch Window to find the bug. Fix it with one line change.

Use the Aggie Advisors data you've already pasted into Excel, or the data
from the Variables module. Your data should have StudentID in column A and
TAMU_GPR in column D.

```vba
Option Explicit
Sub CountHighGPR_Broken()

    Dim HighCount As Integer

    Sheets("Sheet1").Select
    Range("A2").Select

    Do Until ActiveCell = ""

        If ActiveCell > 3.5 Then
            HighCount = HighCount + 1
        End If

        ActiveCell.Offset(1, 0).Select

    Loop

    MsgBox "Students with GPR above 3.5: " & HighCount

End Sub
```

**Hint (show/hide):**
Add a Watch for both `ActiveCell` and `ActiveCell.Offset(0, 3)` before
stepping through. Look at what each one actually contains when the IF
evaluates. One of them is the GPR — which one?

**Solution (show/hide):**
The bug is `If ActiveCell > 3.5` — this checks column A (StudentID),
not column D (GPR). A 9-digit student ID will never be greater than 3.5.

Fix: `If ActiveCell.Offset(0, 3) > 3.5 Then`

Expected result with Aggie Advisors data: Students with GPR above 3.5: 16

---

## AGGIE ADVISORS CHALLENGE SECTION (id="challenge")

### h2: Exam Challenge — Debug the Aggie Advisors Macro

**Intro paragraph:**
This is exam level. The macro below is a broken version of the AddNewStudents
macro from the Aggie Advisors project. It runs without an error but produces
wrong results — it adds the wrong number of students and the count in the
final MsgBox is incorrect.

There are two bugs. Find both using F8 and the Watch Window. Fix each with
the minimum change possible — one line per bug. No hints.

```vba
Option Explicit
Sub AddNewStudents_Broken()

    Dim NewGroup       As Integer
    Dim NumberAccepted As Integer
    Dim StudentID      As Long
    Dim GPR            As Double
    Dim LastName       As String
    Dim FirstName      As String

    Sheets("Applicant Information").Select
    Range("A2").Select

    If ActiveCell = "" Then
        MsgBox "No applicants found."
        Exit Sub
    End If

    NewGroup = InputBox("Enter the new group number:")

    Do Until ActiveCell = ""

        If ActiveCell.Offset(0, 6) = "Accept" Then

            StudentID = ActiveCell
            LastName  = ActiveCell.Offset(0, 1)
            FirstName = ActiveCell.Offset(0, 2)
            GPR       = ActiveCell.Offset(0, 3)

            Sheets("Student Information").Select

            ActiveCell           = StudentID
            ActiveCell.Offset(0, 1) = NewGroup
            ActiveCell.Offset(0, 2) = "U"
            ActiveCell.Offset(0, 3) = LastName
            ActiveCell.Offset(0, 4) = FirstName
            ActiveCell.Offset(0, 5) = GPR

            NumberAccepted = NumberAccepted + 1

            ActiveCell.Offset(1, 0).Select
            Sheets("Applicant Information").Select

        End If

        ActiveCell.Offset(1, 0).Select

    Loop

    MsgBox NumberAccepted & " students added for Group " & NewGroup

End Sub
```

**What the correct output should be:**
- 20 students added (all Accept decisions from the 30-record dataset)
- MsgBox: "20 students added for Group [X]"

**The two bugs:**
Bug 1 — Wrong column offset in the IF condition:
`ActiveCell.Offset(0, 6)` checks column G (index 6 from A).
FinalDecision is in column H — that's Offset(0, 7).
Fix: `If ActiveCell.Offset(0, 7) = "Accept" Then`

Bug 2 — Student Information navigation is wrong:
After switching to Student Information and populating fields,
`ActiveCell.Offset(1, 0).Select` moves down one row on Student Information.
But the macro hasn't navigated to the correct next empty row first —
it's writing over existing students.
The move to the next empty row needs to happen before populating,
using `Selection.End(xlDown).Select` then `ActiveCell.Offset(1, 0).Select`.
(Note: students should find Bug 1 with F8 first since it's more obvious —
Bug 2 may require stepping through a few accepted records to notice.)

**Link:** See the full Aggie Advisors project context →
`href="/src/modules/practice-project.html#module-8"`

### Module 9 — Pseudocode
[To be written before pseudocode.html is built]

### Practice Project Page
[To be written before practice-project.html is built]