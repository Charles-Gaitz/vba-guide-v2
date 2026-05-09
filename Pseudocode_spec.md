# Pseudocode — Content Spec
# Module 9 of 9
# File: src/modules/pseudocode.html
# Prev: /src/modules/debugging.html (F8 Debugging Practice)
# Next: none (last module)
# Note: module-nav next slot should be Practice Project link

---

## Canvas prerequisite (.box-reminder):
"REMINDER: To fully understand Pseudocode, you should have already watched
the Pseudocode Video in Canvas and followed along with the Macro Demo file.
This practice will build upon that foundation."

---

## CONCEPT SECTION (id="concept")
h2 heading: "Pseudocode in VBA"

### Opening — What Pseudocode Is

**Paragraph 1:**
Pseudocode is plain-language logic written before you write any code.
It's not VBA — it has no syntax rules, it won't compile, and Excel
can't run it. Its only job is to help you think through the logic of
a macro before you start recording or typing. A well-written pseudocode
is like an outline for an essay: if the outline is clear, the writing
goes much faster and stays on track.

**Paragraph 2:**
Professor Sanders requires pseudocode as the first step of the macro
project for a reason. Students who skip it and go straight to code
spend far more time debugging than students who plan first. Pseudocode
forces you to identify what programming concepts you need — IF statements,
variables, loops — before you're in the middle of writing code and
second-guessing yourself.

**course-tip (concept):**
"The pseudocode assignment is graded on communication, not syntax.
Ask yourself: if someone who doesn't know VBA read this, would they
understand what the macro does? If yes, it's good pseudocode."

---

#### Sanders Pseudocode Format

**Paragraph 1:**
Professor Sanders uses a specific pseudocode format in this course.
Keywords are CAPITALIZED. Logic inside IF blocks and loops is indented.
Every IF ends with END IF. Every loop starts with a loop instruction
and ends with ENDLOOP or NEXT. Variable names and field names are
used consistently throughout — not "the student" in one place and
"applicant" in another.

**.syntax-box:**
```
DEFINE Variables: VariableName1, VariableName2

IF condition THEN
    PERFORM action
    DISPLAY result
ELSE
    DISPLAY alternate message
END IF

SELECT first record
DO UNTIL End_of_File
    PERFORM action on current record
    MOVE to next record
ENDLOOP
```

**Introduction sentence before code:**
"This is the pseudocode structure from the Project Demo — the logic
for adding accepted students to the roster:"

```
DEFINE Variables: NewGroup, NumberAccepted, UIN, GPR

IF Applicant Records is Empty THEN
    DISPLAY Message "No applicants found"
ELSE
    PROMPT user for NewGroup
    DISPLAY NewGroup on Valid Values
    COPY Applicant Information sheet as backup

    SELECT first Applicant record

    DO UNTIL End_of_File
        IF FinalDecision = Accept THEN
            POPULATE variables for UIN, GPR
            SELECT Student Information
            DISPLAY UIN, GPR, NewGroup, TrackCode
            ADD 1 to NumberAccepted
            MOVE to next Student row
            SELECT Applicant Information
        END IF
        MOVE to next Applicant
    ENDLOOP

    REFRESH reports
    DISPLAY "X Students Added for Group Y"
END IF
```

---

#### Pre-Pseudocode Questions

**Paragraph 1:**
Before writing pseudocode, ask yourself four questions about the problem.
The answers tell you exactly which programming concepts your macro needs.
These questions come directly from Professor Sanders' Macro Handout and
are the starting point for every macro you plan.

**.syntax-box:**
```
1. Is there any data you need to get from the user?
   → Yes = InputBox + Variable

2. Does any value change each time you run it?
   → Yes = Variable

3. Is there anything dependent on a condition?
   → Yes = IF Statement

4. Is anything repetitive?
   → Yes = Loop
   → Do you know how many times? Yes = For Next, No = Do Until
```

---

#### From Pseudocode to Code

**Paragraph 1:**
Good pseudocode maps almost directly to VBA structure. A DEFINE line
becomes a Dim statement. A PROMPT line becomes an InputBox. A DO UNTIL
loop becomes exactly that in code. An IF THEN becomes an If/Then/End If.
The logic and the structure are the same — only the syntax changes.

**.syntax-box:**
```
Pseudocode → VBA

DEFINE NewGroup As Integer    →  Dim NewGroup As Integer
PROMPT user for NewGroup      →  NewGroup = InputBox("Enter group number")
IF records empty THEN         →  If ActiveCell = "" Then
    DISPLAY message           →      MsgBox "No records found"
    STOP                      →      Exit Sub
END IF                        →  End If
DO UNTIL End_of_File          →  Do Until ActiveCell = ""
    MOVE to next record       →      ActiveCell.Offset(1, 0).Select
ENDLOOP                       →  Loop
```

---

### QUICK CHECK SECTION (id="quick-check")

**Question 1:**
What is the main purpose of pseudocode?
- A. To run a macro without the VBA Editor
- B. To plan the logic before writing actual code ← CORRECT
- C. To document code after it's been written
- D. To check for syntax errors
**Explanation:** Pseudocode is planning, not programming. It has no
syntax rules and can't run — its only purpose is to think through
the logic before you start writing VBA.

**Question 2:**
In Sanders pseudocode format, how should keywords be written?
- A. In lowercase
- B. In italics
- C. In CAPITALS ← CORRECT
- D. In quotes
**Explanation:** CAPITALIZED keywords like DEFINE, DISPLAY, POPULATE,
DO UNTIL, and ENDLOOP make the structure of the pseudocode visually
clear and easy to read.

**Question 3:**
You need a macro that asks for a date, uses it in a calculation, and
runs differently depending on whether the result is above or below a
threshold. Which concepts do you need?
- A. Just a loop
- B. Variable and IF statement ← CORRECT
- C. Just an InputBox
- D. Loop and IF statement
**Explanation:** The date is user input → Variable. The different
behavior based on threshold → IF statement. No loop is needed since
you're not processing multiple records.

**Question 4:**
Which pseudocode line correctly ends an IF block?
- A. STOP IF
- B. CLOSE IF
- C. END IF ← CORRECT
- D. ENDIF (no space)
**Explanation:** Sanders format uses END IF (two words) to close every
IF block. Every IF must have a matching END IF.

**Question 5:**
You're writing pseudocode for a macro that processes every student in
a table but you don't know how many students there are. Which loop
instruction is correct?
- A. FOR Count = 1 TO NumberOfStudents
- B. DO UNTIL End_of_File ← CORRECT
- C. REPEAT UNTIL Done
- D. WHILE records remain
**Explanation:** When you don't know the count, use DO UNTIL with an
end-of-file condition. FOR loops are for when you know the exact count.

**course-tip after quick check:**
"The pseudocode assignment is due before you write any code — that's
intentional. If your pseudocode is solid, writing the VBA is mostly
just translation. If you're stuck on the code, go back to the pseudocode."

---

### EASY WINS SECTION (id="easy-wins")

#### Exercise 1 — Answer the Four Questions (SIMPLE FORMAT)
**Difficulty:** Observation

Read this task description and answer the four pre-pseudocode questions:

*"Write a macro that loops through all students in the Applicant
Information sheet and counts those whose TAMU GPR is above 3.5.
Ask the user what GPR threshold to use instead of hardcoding 3.5.
Display the count when done."*

Answer each question:
1. Is there data to get from the user?
2. Does any value change each time you run it?
3. Is there anything conditional?
4. Is anything repetitive? If yes, do you know how many times?

**Hint:** Work through each question one at a time before looking
at the solution.

**Solution:**
1. Yes — the GPR threshold → InputBox + Variable (Double)
2. Yes — the threshold changes each run → Variable
3. Yes — only count students above the threshold → IF statement
4. Yes — process every student → Loop. Don't know how many → Do Until

Concepts needed: Variable, InputBox, IF statement, Do Until loop.

---

#### Exercise 2 — Write the Pseudocode (STEPS FORMAT)
**Difficulty:** Guided

Using the task from Exercise 1, write the pseudocode using Sanders format.

**Step 1 — Define your variables**
Start with DEFINE and list the variables you identified in Exercise 1.
Note the data type in parentheses after each variable name — this is
the Sanders pseudocode convention, not a VBA comment.

```
DEFINE Variables: GPRThreshold (Double), HighGPRCount (Integer)
```

**Step 2 — Get user input**
Add a PROMPT line for the threshold:
```
PROMPT user for GPRThreshold
```

**Step 3 — Add the loop and IF**
```
SELECT first student record

DO UNTIL End_of_File
    IF TAMU_GPR > GPRThreshold THEN
        ADD 1 to HighGPRCount
    END IF
    MOVE to next student
ENDLOOP
```

**Step 4 — Display the result**
```
DISPLAY "Students above " GPRThreshold ": " HighGPRCount
```

**Complete Pseudocode:**
```
DEFINE Variables: GPRThreshold (Double), HighGPRCount (Integer)

PROMPT user for GPRThreshold

SELECT first student record

DO UNTIL End_of_File
    IF TAMU_GPR > GPRThreshold THEN
        ADD 1 to HighGPRCount
    END IF
    MOVE to next student
ENDLOOP

DISPLAY "Students above " GPRThreshold ": " HighGPRCount
```

---

### PRACTICE PROBLEM SECTION (id="practice-problem")

#### No data table needed for this module.
The practice problem is a planning exercise, not a coding exercise.
Use .sample-data-exercise directly with no .data-table-section.
Do NOT include an .exercise-hint component — students are writing
pseudocode, not code, so there is nothing to hint about syntax.

#### module-nav next slot:
href="/src/modules/practice-project.html" text "Practice Project →"

#### Practice Problem — Pseudocode the Guard Check Macro
Write pseudocode for the guard check macro from Module 2 using
Sanders format. The macro should:
- Check if the Applicant Information sheet has records
- If empty: display a message and stop
- If not empty: ask for a group number, then display a ready message

**Requirements:**
- Use DEFINE for any variables
- Use CAPITALS for all keywords
- Indent logic inside IF blocks
- End every IF with END IF

**Expected pseudocode:**
```
DEFINE Variables: NewGroup (Integer)

SELECT Applicant Information sheet

IF A2 is empty THEN
    DISPLAY "No applicants found"
    STOP
ELSE
    PROMPT user for NewGroup
    DISPLAY "Group " NewGroup " is ready to process"
END IF
```

**Link:** See this in the Aggie Advisors project →
`href="/src/modules/practice-project.html#module-9"`

---

### EXAM CHALLENGE SECTION (id="challenge")

**Title:** Pseudocode the Full Aggie Advisors Macro
**No hints. Exam level.**

Write complete pseudocode for the full AddNewStudents macro from
the Aggie Advisors project. Your pseudocode must cover all of the
following using Sanders format:

1. Variable definitions
2. Guard check for empty applicant list
3. User prompt for group number
4. Copy Applicant Information as backup
5. The main processing loop — reading variables, adding to Student
   Information, counting accepted students
6. Refresh reports
7. Completion message

Your pseudocode should be detailed enough that someone who doesn't
know VBA could understand exactly what the macro does, step by step.

Refer to the Aggie Advisors practice project for the full macro
context if needed.

**Link:** See the full Aggie Advisors project →
`href="/src/modules/practice-project.html#module-9"`