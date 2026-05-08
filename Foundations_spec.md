# Macro Foundations — Content Spec
# Module 1 of 9
# File: src/modules/foundations.html
# Prev: none (first module)
# Next: /src/modules/programming-concepts.html (Adding Programming Concepts)

---

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