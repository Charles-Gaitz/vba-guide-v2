# VBA Practice for ACCT 628 — Content Specification

## Ground Rules (apply to every module)

- All VBA code examples must come directly from this file. Never invent examples.
- All terminology matches Professor Sanders exactly:
  - "DEFINE, POPULATE, USE" — the three steps of variables
  - "recorded" vs "typed" — how to distinguish code types
  - "hardcoded" — values that should be replaced with variables
  - "ActiveCell", "Offset", "Do Until" — primary navigation and loop patterns
- Pseudocode: keywords CAPITALIZED, indented inside loops/IF, END IF, ENDLOOP
- The Sanders approach is always the reference. AI code appears ONLY in .ai-compare.
- .course-tip voice: peer-to-peer, CE to fellow students, short and direct.
- Concept sections: minimum 2 paragraphs of prose per concept BEFORE any code appears.
  Introduce what the code shows before showing it. Never drop a code block without
  a preceding introduction sentence.
- Exercise format per module is specified explicitly below — steps or simple.

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
When you write a macro, it runs from top to bottom — one line, then the next, then
the next. That works fine for a handful of actions. But what if you need to do the
same thing to 30 student records? Or 300? Writing the same lines of code 30 times
isn't just tedious — it means if you need to change something, you have to change it
30 times. Loops solve this. A loop lets you write the logic once and repeat it
automatically for as many records as you have.

**Paragraph 2:**
VBA gives you three loop structures: For Next, Do While, and Do Until. They all
repeat a block of code, but they differ in how they decide when to stop. Choosing
the right one for your situation makes your macro clearer and easier to debug.
For this course, the Do Until loop is the one you'll use most — especially for
processing rows of data where you don't know in advance how many records there are.

**course-tip (concept):**
"The 'move to next record' line at the bottom of the loop is the one students
forget most often. If you forget it on a Deny record, the loop runs forever on
that row. Write the move line before you write anything else inside the loop."

---

#### For Next Loop

**Paragraph 1:**
The For Next loop is the right choice when you know exactly how many times the loop
needs to run. It uses a counter variable that starts at a number you specify, runs
the code inside, then increases the counter by 1, and checks if it's reached the
end value. If not, it runs again. If so, it stops.

**Paragraph 2:**
Think of it like telling Excel: "Do this starting from step 1, and keep going until
you've done it 4 times." The counter variable tracks where you are. You can use any
variable name for the counter — `i` is common by convention — and you can use
variables for the start and end values instead of hardcoded numbers.

**Introduction sentence before code:**
"Here is what a For Next loop looks like — this example puts the numbers 1 through
4 into adjacent columns:"

```vba
Sub ForNextLoop()
    ' Makes cell A1 on Sheet2 the active cell
    Sheets("Sheet2").Select
    Range("A1").Select
    Cells.ClearContents   ' Deletes all the cells on the worksheet

    ' Goes through the loop 4 times and inserts the value of i into the selected cell
    ' Moves to the right i number of cells
    For i = 1 To 4
        ActiveCell.Value = i
        ActiveCell.Offset(0, i).Select
    Next i
End Sub
```

---

#### Do While Loop

**Paragraph 1:**
The Do While loop runs as long as a condition is true. Before each pass through
the loop, VBA checks the condition. If it's true, the code inside runs. If it's
false, the loop stops and VBA moves on to whatever comes after it. Importantly,
if the condition is already false when the loop starts, the code inside never
runs at all.

**Paragraph 2:**
The key word here is "while" — the loop keeps going while something is happening.
In the example below, the loop keeps running while `myNum` is greater than 10.
Each time through, `myNum` gets smaller by 1. When it hits 10, the condition
becomes false and the loop stops. Notice that `myNum` must change inside the loop —
if it never changed, the condition would stay true forever and the loop would
never stop.

**Introduction sentence before code:**
"Here is a Do While loop that counts how many times it subtracts 1 from a number
before it reaches 10:"

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
The Do Until loop is the mirror image of Do While. Instead of running while
a condition is true, it runs until a condition becomes true — which means it keeps
going while the condition is false. This might sound confusing at first, but
it becomes natural once you see it in context. Think of it as: "keep doing this
until you're done."

**Paragraph 2:**
For processing rows of data in Excel, Do Until is almost always the right choice.
The condition you check is whether the active cell is blank — `Do Until ActiveCell = ""`.
As long as there's data in the current cell, the loop continues. When you reach an
empty cell, the loop stops. This works regardless of how many records are in your
spreadsheet, which makes your macro adaptable instead of hardcoded.

**Introduction sentence before code:**
"Here is a Do Until loop using the same counting example — notice how it reads
almost like the opposite of Do While:"

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
In this course, almost every time you use a loop you'll be using it to process
rows of data in Excel — going through a list of students, applicants, or records
and doing something with each one. This pattern is so common that it's worth
memorizing as a single unit. Professor Sanders calls it the "Processing All Records"
pattern, and it looks the same every time you use it.

**Paragraph 2:**
The pattern has three parts: navigate to the first record, loop until you hit a
blank cell, and at the very end of the loop body, move down one row. That last
part — moving down — is what makes the loop progress. Without it, the loop checks
the same cell forever. With it, the loop moves through every row until it runs out
of data and hits a blank.

**Introduction sentence before pseudocode:**
"In pseudocode, the pattern looks like this every time:"

```
SELECT first record
DO UNTIL ActiveCell = ""
    PERFORM action on current row
    MOVE to next record
ENDLOOP
```

**Introduction sentence before real project code:**
"And here is what it looks like in a real macro — this is the actual loop from
the Project Demo that processes student applicants:"

```vba
' SELECT Starting Location for Applicant Information (first record)
Sheets("Applicant Information Group " & NewGroup & "").Select
Range("A2").Select

Do Until ActiveCell = ""

    If ActiveCell.Offset(0, 10) = Range("accept") Then

        ' POPULATE variables for UIN, GPR, Gender, Birthdate
        UIN = ActiveCell
        GPR = ActiveCell.Offset(0, 1)

        ' SELECT Student Information and DISPLAY values
        Sheets("Student Information").Select
        ActiveCell = UIN
        ' ... more population ...

        ' MOVE to next Student row
        ActiveCell.Offset(1, 0).Range("PPAData[[#Headers],[UIN]]").Select

        ' RETURN to Applicant Information
        Sheets("Applicant Information Group " & NewGroup & "").Select

    End If

    ' ALWAYS move to next applicant — whether accepted or not
    ActiveCell.Offset(1, 0).Range("ApplicantData[[#Headers],[UIN]]").Select

Loop
```

---

#### Cautions: Endless Loops

**Paragraph 1:**
An endless loop is what happens when the condition never becomes true — the loop
runs forever and Excel freezes. This almost always happens for one of three reasons:
the condition doesn't involve a variable, the variable never changes inside the loop,
or the "move to next record" line is inside an IF block instead of outside it.

**Bullet points:**
- The condition MUST involve at least one variable
- That variable MUST change value somewhere inside the loop body
- The "move to next record" line must be OUTSIDE any IF block — it should
  always execute, whether the IF condition is true or false
- If you get stuck in an endless loop: press **Ctrl+Break** to stop it

---

#### AI Compare Panel

**Sanders panel label:** ✅ Sanders Approach
**Sanders panel h4:** Simple Do Until — mirrors how you navigate in Excel

```vba
' Sanders approach — easy to follow with F8
Sheets("Student Information").Select
Range("A2").Select

Do Until ActiveCell = ""
    ' process current row
    ActiveCell.Offset(1, 0).Select  ' move to next record
Loop
```

**AI panel label:** ⚠️ Typical AI / Google Result
**AI panel h4:** For Each with object variables

```vba
' AI approach — uses concepts this course doesn't cover
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
Both loops process every record — but they don't look the same. The Sanders
approach mirrors how you navigate in Excel: start at A2, loop until blank, move
down one row. You can watch it run with F8 and see exactly where the cursor is
at every step.

The AI version works, but it uses concepts this course doesn't cover — object
variables, For Each, finding the last row with End(xlUp). If you use code like
this without understanding it, debugging with F8 will be confusing and the
exam won't go well.

---

### QUICK CHECK SECTION

**Format:** Multiple choice. Lock on click. Reveal correct answer immediately.

**Question 1:**
You have a macro that needs to process every student record in a spreadsheet.
You don't know how many students there are. Which loop should you use?
- A. For Next Loop
- B. Do Until Loop ← CORRECT
- C. Do While Loop
- D. It doesn't matter, they all work the same
**Explanation after reveal:** Do Until is the right choice when you don't know
how many records there are. You loop until you hit a blank cell — this works
for 10 records or 10,000 without changing a single line of code.

**Question 2:**
In a Do Until loop, the code inside the loop runs when the condition is _____.
- A. True
- B. False ← CORRECT
- C. Either true or false
- D. Only on the first pass
**Explanation after reveal:** Do Until runs while the condition is false, and
stops when it becomes true. So `Do Until ActiveCell = ""` keeps running while
the cell is NOT blank, and stops the moment it hits a blank cell.

**Question 3:**
What happens if you forget the `ActiveCell.Offset(1, 0).Select` line at the
bottom of your Do Until loop?
- A. The loop skips every other row
- B. VBA throws an error and stops
- C. The macro runs forever on the same cell ← CORRECT
- D. The loop only runs once
**Explanation after reveal:** Without moving to the next row, the loop checks
the same cell over and over. The condition never becomes true, so it never stops.
Press Ctrl+Break if you get stuck.

**Question 4:**
Which of the following is the correct pattern for processing all records
in an Excel table using a Do Until loop?
- A. `Do Until ActiveCell = 0`
- B. `Do Until ActiveCell = "done"`
- C. `Do Until ActiveCell = ""` ← CORRECT
- D. `Do Until ActiveCell.Row = 100`
**Explanation after reveal:** Checking for a blank cell (`""`) is the standard
pattern. It works regardless of how many records are in the table, which makes
your macro adaptable instead of hardcoded to a specific row number.

**Question 5:**
In a For Next loop with `For i = 1 To 4`, how many times does the loop run?
- A. 3 times
- B. 4 times ← CORRECT
- C. 5 times
- D. It depends on the data
**Explanation after reveal:** The loop runs for i = 1, 2, 3, and 4 — four
times total. The counter starts at the start value and runs through the end
value inclusive.

**course-tip after quick check:**
"Question 2 trips people up every semester. Do Until sounds like it runs
until something is true — but the code inside runs while the condition is
false. Draw it out: the loop asks 'am I done yet?' and keeps going as long
as the answer is no."

---

### EASY WINS SECTION

#### Exercise 1 — For Next Loop (STEPS FORMAT)
**Title:** Write Your First For Next Loop
**Difficulty badge:** Guided
**Description:** Write a macro that uses a For Next loop to put the numbers
1 through 5 into cells A1 through A5, one number per cell.

**Step 1 — Open the VBA Editor**
Press Alt+F11 to open the VBA Editor. Go to Insert → Module. You should
see a blank white area. This is where you write your code.

Type this shell to start:
```vba
Option Explicit
Sub NumberCells()

End Sub
```
`Option Explicit` forces you to declare variables. Always include it.

**Step 2 — Navigate to your starting cell**
Before the loop starts, you need to tell Excel where to begin. Add these
two lines inside your Sub, before the loop:
```vba
Sheets("Sheet1").Select
Range("A1").Select
```
This selects Sheet1 and moves to cell A1 — your starting point.

**Step 3 — Write the For Next loop**
Now add the loop. It should count from 1 to 5 and put the current
counter value into the active cell:
```vba
For i = 1 To 5
    ActiveCell.Value = i
    ActiveCell.Offset(1, 0).Select
Next i
```
`ActiveCell.Value = i` puts the current number into the cell.
`ActiveCell.Offset(1, 0).Select` moves down one row.
`Next i` increases i by 1 and loops back.

**Step 4 — Run it**
Press F5 or click the green Play button. Switch to Excel and check
column A. You should see 1, 2, 3, 4, 5 in cells A1 through A5.

**View Complete Code:**
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

#### Exercise 2 — What Happens When the Condition Is Already Met? (SIMPLE FORMAT)
**Title:** Test the Do Until Condition
**Difficulty badge:** Observation
**Description:** This is a quick experiment, not a build exercise.

Go back to the `ChkFirstUntil` example from the concept section above.
Change `myNum = 20` to `myNum = 9`. Run the macro.

What happens? Why?

**Hint:** Think about what the condition `Do Until myNum = 10` checks
before the loop even starts.

**Solution:** Nothing happens — the MsgBox shows 0 repetitions.
Because `myNum` starts at 9, which is already less than 10, the condition
`myNum = 10` is already false... wait — actually `myNum = 9` means `myNum = 10`
is false, so `Do Until myNum = 10` starts with a false condition.
That means the loop DOES run. But wait — `myNum - 1` makes it 8, 7, 6...
it never equals 10, so the loop runs forever.

Actually: set `myNum = 10` to see the loop not run at all (condition already
true at start). Set `myNum = 9` to see an endless loop (condition can never
become true). Press Ctrl+Break to stop it.

**Key takeaway:** Do Until won't run if the condition is already true when
it starts. And if the condition can never become true, it runs forever.

---

#### Exercise 3 — Do Until Counter (STEPS FORMAT)
**Title:** Count Non-Blank Cells with Do Until
**Difficulty badge:** Guided
**Description:** Write a macro that starts at cell A1, uses a Do Until loop
to move down the column, and counts every non-blank cell it finds. Display
the count in a MsgBox when done.

**Step 1 — Set up your macro**
Open the VBA Editor (Alt+F11). Add a new module (Insert → Module) or
use the same one. Create this shell:
```vba
Option Explicit
Sub CountCells()

End Sub
```

**Step 2 — Declare your variable**
You need one variable to keep a running total of non-blank cells.
Add this inside your Sub:
```vba
Dim CellCount As Integer
```
`Integer` is fine here — we won't have more than 32,000 rows.

**Step 3 — Navigate to your starting cell**
Add these lines to start at A1:
```vba
Sheets("Sheet1").Select
Range("A1").Select
```

**Step 4 — Write the Do Until loop**
Now add the loop. It should keep going until it hits a blank cell,
add 1 to the counter on each pass, and move down one row:
```vba
Do Until ActiveCell = ""
    CellCount = CellCount + 1
    ActiveCell.Offset(1, 0).Select
Loop
```
Notice `CellCount = CellCount + 1` — this takes the current value of
CellCount, adds 1, and stores the result back into CellCount. This is
how you accumulate a total.

**Step 5 — Display the result**
After the loop ends, display the count:
```vba
MsgBox "Non-blank cells found: " & CellCount
```

**View Complete Code:**
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
**Expected result:** MsgBox shows the number of filled cells in column A.

---

### SAMPLE DATA SECTION

#### Data Table (Aggie Advisors — 30 records)
**Collapsed by default. Copy button copies as TSV.**
**Instructions above table:** "Copy this data and paste it into cell A1
of a new Excel worksheet. Use Ctrl+V to paste — Excel will automatically
split the columns. This is your practice dataset for the exercises below."

The 30-record dataset is the full Applicant Information table from
PRACTICE_PROJECT.md — StudentID, LastName, FirstName, TAMU_GPR,
Grade229, Grade230, Grade327, FinalDecision.

#### Sample Data Exercise
**Title:** Count Accepted Applicants
**Description:** Using the Aggie Advisors data you just pasted into Excel,
write a macro that loops through all 30 applicant records and counts only
the students whose FinalDecision column says "Accept". Display the result
in a MsgBox.

**What your macro needs to do:**
- Start at cell A2 (row 1 is the header)
- Use Do Until ActiveCell = "" to loop through all records
- Inside the loop, check if the FinalDecision column = "Accept"
  (FinalDecision is in column H — that's Offset(0, 7) from column A)
- Add 1 to a counter variable each time you find an "Accept"
- Move to the next row at the bottom of the loop — every time,
  whether the decision was Accept or not
- After the loop, display: "Accepted applicants: " & your counter

**Expected answer:** Accepted applicants: 20

**Hint:** Your IF statement should look like this:
```vba
If ActiveCell.Offset(0, 7) = "Accept" Then
    AcceptCount = AcceptCount + 1
End If
```
The move line goes AFTER the End If, not inside it.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-4"`

---

### EXAM CHALLENGE SECTION

**Title:** Accepted and Denied GPR Averages
**No hints. No steps. This is exam level.**

Using the same Aggie Advisors dataset, write a macro that:
1. Loops through all 30 applicant records using Do Until with blank cell check
2. Separately accumulates total GPR for accepted students and denied students
3. Counts accepted and denied students separately
4. After the loop, calculates the average GPR for each group
5. Displays both averages in a single MsgBox:
   "Accepted avg GPR: X.XX | Denied avg GPR: X.XX"

Must use Option Explicit. Must declare all variables with correct data types.
Must use a Named Range or the literal string "Accept" consistently —
do not mix approaches.

**Expected answers:**
- Accepted: 20 students, average GPR 3.7175
- Denied: 10 students, average GPR 3.0014

If your numbers don't match, use F8 and the Watch Window to step through
the first few records and check what your IF condition is actually comparing.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-4"`

---

## Modules 1, 2, 3, 5, 6, 7 — Practice Section Format

The practice section format for each remaining module follows the same
structure as Loops above. Each module's content spec will be written
here before that module's task file is created. Do not build any module
page without the full practice content being specified in this file first.

Placeholder sections below will be fleshed out as each module is built
in sequence.

### Module 1 — Macro Foundations
[Practice content to be written before foundations.html is built]

### Module 2 — Adding Programming Concepts
[Practice content to be written before programming-concepts.html is built]

### Module 3 — Variables
[Practice content to be written before variables.html is built]

### Module 5 — Relative vs Absolute References
[Practice content to be written before references.html is built]

### Module 6 — Filters & Shortcut Keys
[Practice content to be written before filters.html is built]

### Module 7 — F8 Debugging
[Practice content to be written before debugging.html is built]