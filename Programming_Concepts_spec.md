# Adding Programming Concepts — Content Spec
# Module 2 of 9
# File: src/modules/programming-concepts.html
# Prev: /src/modules/foundations.html (Macro Foundations)
# Next: /src/modules/variables.html (Variables)

---

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

**Fix:** Correct the spelling to MyValue = 10. Run again — no error.

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