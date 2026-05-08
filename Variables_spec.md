# Variables Content Spec — Append to CONTENT_SPEC.md
# Replace the placeholder line:
# "### Module 3 — Variables"
# "[To be written before variables.html is built]"
# with the full content below:

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