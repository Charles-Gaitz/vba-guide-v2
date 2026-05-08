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
[To be written before foundations.html is built]

### Module 2 — Adding Programming Concepts
[To be written before programming-concepts.html is built]

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
[To be written before debugging.html is built]

### Module 9 — Pseudocode
[To be written before pseudocode.html is built]

### Practice Project Page
[To be written before practice-project.html is built]