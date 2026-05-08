# F8 Debugging Practice — Content Spec
# Module 8 of 9
# File: src/modules/debugging.html
# Prev: /src/modules/filters.html (Filters & Shortcut Keys)
# Next: /src/modules/pseudocode.html (Pseudocode)

---

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