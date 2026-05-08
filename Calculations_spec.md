# Calculations and Dates — Content Spec
# Module 5 of 9
# File: src/modules/calculations.html
# Prev: /src/modules/loops.html (Loops)
# Next: /src/modules/references.html (Relative vs Absolute References)

---

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Calculations and Dates, you should have already
watched the Calculations and Dates Video in Canvas and followed along with the
Macro Demo file. This practice will build upon that foundation."

---

## CONCEPT SECTION (id="concept")

### Opening — Calculations and Dates in VBA

**Paragraph 1:**
VBA can perform calculations three ways: store the result in a variable,
place a formula in a cell so Excel calculates it, or place the result of
the calculation directly in a cell. Each approach has its place. Variables
are used when you need the value mid-macro. Formulas in cells let Excel
recalculate automatically. Results in cells are faster and don't depend
on Excel's calculation engine.

**Paragraph 2:**
For this course, the most common pattern is calculating into a variable
and then displaying or storing the result. Functions like Count, Sum, and
Average are available through Application.WorksheetFunction and work
exactly like their Excel equivalents.

**course-tip (concept):**
"If a function works in Excel, it almost certainly works in VBA through
Application.WorksheetFunction. When in doubt, try it — if it doesn't work
without the prefix, add it."

---

#### Basic Math and WorksheetFunction

**Paragraph 1:**
Basic arithmetic in VBA uses the same operators as Excel: +, -, *, /.
You can calculate directly into a variable or into a cell. To use Excel
functions like Count, Sum, or Average in VBA, prefix them with
Application.WorksheetFunction. Some functions work without the prefix
— but when in doubt, include it.

**.syntax-box:**
```
' Basic math into a variable
AverageGPR = TotalGPR / StudentCount

' WorksheetFunction into a variable
NoRecords = Application.WorksheetFunction.Count(Range("A:A"))
TotalSales = Application.WorksheetFunction.Sum(Range("H:H"))

' Place formula in a cell (Excel calculates)
Range("K2") = "=Count(A:A)"

' Place result in a cell (VBA calculates)
Range("K3") = Application.WorksheetFunction.Count(Range("A:A"))
```

**Introduction sentence before code:**
"This example from the Macro Handout shows all three approaches for Count:"

```vba
Dim NoRecords As Integer

' Places the actual count formula in the cell
Range("K2") = "=Count(A:A)"

' Places the result into a cell
Range("K3") = Application.WorksheetFunction.Count(Range("A:A"))

' Places the result into a variable
NoRecords = Application.WorksheetFunction.Count(Range("A:A"))
```

---

#### Accumulating a Total in a Loop

**Paragraph 1:**
When you need a running total across multiple records, you accumulate it
inside a loop using a variable. Start the variable at 0 before the loop,
add to it on each pass, then use or display it after the loop ends.
This is the same pattern whether you're summing sales, counting records,
or calculating an average.

**.syntax-box:**
```
Dim CumulativeTotal As Double
CumulativeTotal = 0              ' reset before loop

Do Until ActiveCell = ""
    CumulativeTotal = CumulativeTotal + ActiveCell.Offset(0, 5)
    ActiveCell.Offset(1, 0).Select
Loop

MsgBox "Total: " & CumulativeTotal
```

**Introduction sentence before code:**
"This is the CumulativeTotal example from the Macro Handout:"

```vba
Dim CumulativeTotal As Double

Range("A2").Select
CumulativeTotal = 0

Do Until ActiveCell = ""
    CumulativeTotal = CumulativeTotal + ActiveCell.Offset(0, 5)
    ActiveCell.Offset(1, 0).Select
Loop

MsgBox "The cumulative total is " & _
       Application.WorksheetFunction.Dollar(CumulativeTotal, 0)
```

---

#### Formatting Numbers and Dates

**Paragraph 1:**
Formatting affects what is displayed in a cell — it does not change what
is stored in a variable. A Double variable holding 3.756 stays 3.756
regardless of how the cell displays it. The Dollar function formats a
number as currency for display in a MsgBox. For worksheet cells, use
the NumberFormat property.

**.syntax-box:**
```
' Format a cell as currency
Range("H2").NumberFormat = "$#,##0.00"

' Format a cell as percentage with 2 decimal places
Range("B2").NumberFormat = "0.00%"

' Dollar function for MsgBox display only
MsgBox Application.WorksheetFunction.Dollar(AverageVariable, 2)
```

---

#### Date Functions

**Paragraph 1:**
VBA has built-in date functions for adding time intervals and finding
end-of-month dates. DateAdd takes an interval code, a number, and a
starting date and returns a new date. EoMonth returns the last day of
a month and requires Application.WorksheetFunction. Year, Month, and
Day extract parts of a date and work directly without any prefix.

**.syntax-box:**
```
' Add 1 month to a date
DateAdd("m", 1, StartDate)

' Add 1 day
DateAdd("d", 1, StartDate)

' Add 1 year
DateAdd("yyyy", 1, StartDate)

' Last day of the month
Application.WorksheetFunction.EoMonth(StartDate, 0)

' Extract year, month, day
Year(StartDate)
Month(StartDate)
Day(StartDate)
```

**Introduction sentence before code:**
"These are the date examples from the Macro Handout:"

```vba
Dim StartDate As Date
StartDate = Range("B2")

' Add 1 month
Range("B4") = DateAdd("m", 1, StartDate)

' End of month
Range("B7") = Application.WorksheetFunction.EoMonth(StartDate, 0)

' Compare year to current year
If Year(Range("B7")) = Year(Now()) Then
    MsgBox "You are in the current year"
End If
```

---

### QUICK CHECK SECTION (id="quick-check")

**Question 1:**
You want to count the number of records in column A using VBA and store
the result in a variable. Which line is correct?
- A. `NoRecords = Count(Range("A:A"))`
- B. `NoRecords = Range("A:A").Count`
- C. `NoRecords = Application.WorksheetFunction.Count(Range("A:A"))` ← CORRECT
- D. `NoRecords = "=Count(A:A)"`
**Explanation:** WorksheetFunction functions that store into a variable
require the Application.WorksheetFunction prefix. Option D places a
formula string in a variable, not the result. Option B counts all cells
including empty ones — not what you want.

**Question 2:**
You format a cell as currency using NumberFormat. Your variable GPR still
holds 3.756. What does the variable contain after the formatting?
- A. "$3.76"
- B. 3.756 ← CORRECT
- C. 3.76
- D. The variable is cleared
**Explanation:** Formatting affects display only — it never changes what
is stored in a variable. GPR remains 3.756 as a Double regardless of
how the cell is formatted.

**Question 3:**
What does `DateAdd("m", 3, StartDate)` return?
- A. The date 3 days after StartDate
- B. The date 3 months after StartDate ← CORRECT
- C. The date 3 years after StartDate
- D. The last day of the month 3 months from StartDate
**Explanation:** "m" is the interval code for months. "d" is days,
"yyyy" is years.

**Question 4:**
You need to accumulate a total inside a loop. Before the loop starts,
you should:
- A. Set the variable to 1
- B. Not initialize it — VBA starts variables at 0 automatically
- C. Set the variable to 0 explicitly ← CORRECT
- D. Declare it inside the loop
**Explanation:** While VBA does initialize numeric variables to 0,
explicitly setting it to 0 before the loop is best practice — it makes
your intent clear and prevents bugs if the macro runs more than once
in the same session.

**Question 5:**
Which function requires Application.WorksheetFunction in VBA?
- A. DateAdd
- B. Year
- C. Month
- D. EoMonth ← CORRECT
**Explanation:** DateAdd, Year, Month, and Day are native VBA functions
and work without any prefix. EoMonth is an Excel worksheet function and
requires Application.WorksheetFunction.

**course-tip after quick check:**
"Question 5 comes up on the exam in the form of code that doesn't work —
you'll be asked to fix it. EoMonth without the prefix is a classic error."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Calculate and Display (STEPS FORMAT)
**Difficulty:** Guided

Open the VBA Editor (Alt+F11), insert a module, and build this macro
step by step. Use your Aggie Advisors data — StudentID in column A,
TAMU_GPR in column D.

**Step 1 — Set up your macro shell**
Open the VBA Editor (Alt+F11). Insert a module (Insert → Module). Type:
```vba
Option Explicit
Sub CalculateTotals()
    Dim RecordCount As Integer
    Dim TotalGPR    As Double
End Sub
```
Option Explicit forces variable declarations. The remaining steps fill in
the lines between the Dim statements and End Sub.

**Step 2 — Navigate and calculate**
Add these lines inside your Sub:
```vba
Sheets("Sheet1").Select

RecordCount = Application.WorksheetFunction.Count(Range("A:A"))
TotalGPR    = Application.WorksheetFunction.Sum(Range("D:D"))
```
Count tallies non-blank numeric cells in the column.
Sum adds all values in the column.

**Step 3 — Display the results**
```vba
MsgBox "Records: " & RecordCount & " | Total GPR: " & _
       Application.WorksheetFunction.Dollar(TotalGPR, 3)
```
Dollar formats the number with a dollar sign and 3 decimal places.
Press F5 to run.

**Complete Code (View Complete Code):**
```vba
Option Explicit
Sub CalculateTotals()
    Dim RecordCount As Integer
    Dim TotalGPR    As Double

    Sheets("Sheet1").Select
    RecordCount = Application.WorksheetFunction.Count(Range("A:A"))
    TotalGPR    = Application.WorksheetFunction.Sum(Range("D:D"))

    MsgBox "Records: " & RecordCount & " | Total GPR: " & _
           Application.WorksheetFunction.Dollar(TotalGPR, 3)
End Sub
```
**Expected result with Aggie Advisors data:** Records: 30 | Total GPR: $111.435

---

#### Exercise 2 — Date Experiment (SIMPLE FORMAT)
**Difficulty:** Observation

Type any date into cell B2 of a worksheet. Then run this macro
and observe what each line puts into the cells:

```vba
Option Explicit
Sub DateDemo()
    Dim StartDate As Date
    StartDate = Range("B2")

    Range("C2") = DateAdd("m", 1, StartDate)
    Range("C3") = DateAdd("d", 1, StartDate)
    Range("C4") = DateAdd("yyyy", 1, StartDate)
    Range("C5") = Application.WorksheetFunction.EoMonth(StartDate, 0)
    Range("C6") = Year(StartDate)
End Sub
```

**Hint:** C5 may display as a number instead of a date. Why?

**Solution:** EoMonth returns a date serial number. The cell needs to
be formatted as a date to display correctly. This is the difference
between what VBA stores (a number) and what Excel displays (a date).
Fix: `Range("C5").NumberFormat = "mm/dd/yyyy"`

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### Data Table (Aggie Advisors — 30 records)
Full 30-record dataset from PRACTICE_PROJECT.md.
Columns: StudentID | LastName | FirstName | TAMU_GPR | Grade229 | Grade230 | Grade327 | FinalDecision

#### Practice Problem — Calculate Average GPR of Accepted Students
Using the Aggie Advisors data, write a macro that loops through all
30 records, accumulates the total GPR of accepted students, counts
them, calculates the average, and displays it rounded to 3 decimal places.

**What your macro needs to do:**
- Declare variables: AcceptCount (Integer), TotalGPR (Double), AverageGPR (Double)
- Loop through all records using Do Until blank cell
- Check if FinalDecision = "Accept" (column H, Offset 0,7)
- If Accept: add GPR (column D, Offset 0,3) to TotalGPR and add 1 to AcceptCount
- After loop: calculate AverageGPR = TotalGPR / AcceptCount
- Display: "Accepted: [X] | Avg GPR: [X.XXX]"
  Use Format(AverageGPR, "0.000") to display 3 decimal places without a dollar sign

**Expected result:** Accepted: 20 | Avg GPR: 3.718

**Hint:** Calculate the average AFTER the loop ends, not inside it.
`AverageGPR = TotalGPR / AcceptCount` goes after the Loop line.

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-5"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Calculations and Dates Summary
**No hints. Exam level.**

Using the Aggie Advisors dataset, write a macro that produces a
complete summary in a MsgBox showing:
- Total applicants (all 30)
- Number accepted and their average GPR
- Number denied and their average GPR
- The month and year of today's date

All values must come from variables — no hardcoded numbers.
GPR averages displayed to 3 decimal places using Format(value, "0.000").
Month and year from Month(Now()) and Year(Now()).

**Expected output format:**
"Total: 30 | Accepted: 20 (Avg GPR: 3.718) | Denied: 10 (Avg GPR: 3.001) | Month: 5 Year: 2026"

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-5"`