## 🔍 OVERVIEW

Logical operations in Excel often rely on **conditions** being evaluated as **TRUE or FALSE**. Based on the result, you can return specific values, perform calculations, or control flow.

You can **nest** functions (i.e., use one function inside another) to create powerful, multi-layered logical expressions.

---

## ✅ 1. `IF()` FUNCTION

### Purpose:

Returns one value if a condition is TRUE and another if FALSE.

### Syntax:

```excel
IF(logical_test, value_if_true, value_if_false)
```

### Example:

```excel
=IF(A1>70, "Pass", "Fail")
```

### Nested Example:

```excel
=IF(A1>90, "A", IF(A1>80, "B", IF(A1>70, "C", "Fail")))
```

---

## ✅ 2. `IFS()` FUNCTION

### Purpose:

Tests multiple conditions and returns the result for the first TRUE condition (avoids deep nesting).

### Syntax:

```excel
IFS(condition1, result1, condition2, result2, ...)
```

### Example:

```excel
=IFS(A1>90, "A", A1>80, "B", A1>70, "C", A1<=70, "Fail")
```

---

## ✅ 3. `SWITCH()` FUNCTION

### Purpose:

Compares a value against a list and returns a result for the first match (good for fixed choices).

### Syntax:

```excel
SWITCH(expression, value1, result1, [value2, result2], ..., [default])
```

### Example:

```excel
=SWITCH(A1, "Red", 1, "Green", 2, "Blue", 3, "Unknown")
```

---

## ✅ 4. `SUMIF()`, `AVERAGEIF()`, `COUNTIF()`

### Purpose:

Perform aggregate operations **based on one condition**.

### Syntax:

- `SUMIF(range, criteria, [sum_range])`
- `AVERAGEIF(range, criteria, [average_range])`
- `COUNTIF(range, criteria)`

### Example:

```excel
=SUMIF(B2:B10, ">100", C2:C10)        ' Sum values in C where B > 100
=AVERAGEIF(A2:A10, "<50", B2:B10)     ' Average B where A < 50
=COUNTIF(C2:C10, ">=75")              ' Count how many C values >= 75
```

---

## ✅ 5. `SUMIFS()`, `AVERAGEIFS()`, `COUNTIFS()`

### Purpose:

Perform aggregate operations **based on multiple conditions**.

### Syntax:

- `SUMIFS(sum_range, criteria_range1, criteria1, ...)`
- `AVERAGEIFS(average_range, criteria_range1, criteria1, ...)`
- `COUNTIFS(criteria_range1, criteria1, ...)`

### Example:

```excel
=SUMIFS(C2:C10, A2:A10, "Math", B2:B10, ">80")     ' Sum of scores in Math > 80
=COUNTIFS(D2:D10, ">=75", E2:E10, "Pass")          ' Count records where score >=75 and status is Pass
```

---

## ✅ 6. `MAXIFS()`, `MINIFS()` (Excel 2016+)

### Purpose:

Find the **maximum or minimum** values based on multiple conditions.

### Syntax:

- `MAXIFS(max_range, criteria_range1, criteria1, ...)`
- `MINIFS(min_range, criteria_range1, criteria1, ...)`

### Example:

```excel
=MAXIFS(C2:C10, A2:A10, "Science")   ' Max score in Science
=MINIFS(D2:D10, B2:B10, "<50")       ' Min of D where B < 50
```

---

## ✅ 7. `AND()`, `OR()`, `NOT()`

### Purpose:

Perform logical operations across multiple conditions.

### Syntax:

- `AND(condition1, condition2, ...)`: Returns TRUE if all conditions are TRUE
- `OR(condition1, condition2, ...)`: Returns TRUE if **any** condition is TRUE
- `NOT(condition)`: Reverses the result (TRUE → FALSE, FALSE → TRUE)

### Examples:

#### Combined with IF:

```excel
=IF(AND(A1>70, B1="Pass"), "Promoted", "Retest")
=IF(OR(A1="A", A1="B"), "Top Grades", "Needs Improvement")
=IF(NOT(A1="Fail"), "Continue", "Stop")
```

---

## ✅ 8. `LET()` FUNCTION (Excel 2021 or Microsoft 365)

### Purpose:

Defines variables inside a formula, improving readability and performance.

### Syntax:

```excel
LET(name1, value1, name2, value2, ..., calculation)
```

### Example:

```excel
=LET(score, A1, grade, IF(score>90, "A", IF(score>80, "B", "C")), grade)
```

You can **combine LET with other logical functions**:

```excel
=LET(
  total, SUM(A1:A10),
  avg, AVERAGE(A1:A10),
  IF(avg>total/10, "Above Avg", "Below Avg")
)
```

---

## 🔁 NESTING LOGICAL FUNCTIONS – ADVANCED EXAMPLES

### Example 1: Nested IF + AND

```excel
=IF(AND(A1>50, B1<100), "Valid", "Invalid")
```

### Example 2: IF + COUNTIFS

```excel
=IF(COUNTIFS(A2:A10, "Math", B2:B10, ">80") > 5, "Success", "Review")
```

### Example 3: SUMIFS + IF + NOT

```excel
=IF(NOT(SUMIFS(C2:C10, A2:A10, "Fail")>100), "Under Control", "High Risk")
```

### Example 4: LET + IFS + AVERAGEIFS

```excel
=LET(
  avgScore, AVERAGEIFS(C2:C10, A2:A10, "Science"),
  IFS(avgScore>90, "Excellent", avgScore>70, "Good", TRUE, "Needs Work")
)
```

---

## 🧪 PRACTICE EXERCISE

**Scenario:** You have student data:

| Name | Subject | Score | Status |
| ---- | ------- | ----- | ------ |
| John | Math    | 95    | Pass   |
| Jane | English | 68    | Fail   |
| ...  | ...     | ...   | ...    |

Try creating the following:

1. Count how many students **passed** with a score above 80:

```excel
=COUNTIFS(D2:D100, "Pass", C2:C100, ">80")
```

2. Use `IFS` to assign grades:

```excel
=IFS(C2>=90, "A", C2>=80, "B", C2>=70, "C", TRUE, "F")
```

3. Use `SWITCH` to assign points:

```excel
=SWITCH(B2, "Math", 5, "English", 3, "Science", 4, 0)
```

---

## ✅ SUMMARY TABLE

| Function       | Type        | Use Case                                      |
| -------------- | ----------- | --------------------------------------------- |
| `IF()`         | Logical     | One condition                                 |
| `IFS()`        | Logical     | Multiple conditions (cleaner than nested IFs) |
| `SWITCH()`     | Logical     | Multiple matches based on value               |
| `SUMIF()`      | Conditional | Sum with one condition                        |
| `SUMIFS()`     | Conditional | Sum with multiple conditions                  |
| `COUNTIF()`    | Conditional | Count with one condition                      |
| `COUNTIFS()`   | Conditional | Count with multiple conditions                |
| `AVERAGEIF()`  | Conditional | Average with one condition                    |
| `AVERAGEIFS()` | Conditional | Average with multiple conditions              |
| `MAXIFS()`     | Conditional | Maximum with multiple conditions              |
| `MINIFS()`     | Conditional | Minimum with multiple conditions              |
| `AND()`        | Logical     | All conditions must be TRUE                   |
| `OR()`         | Logical     | At least one condition must be TRUE           |
| `NOT()`        | Logical     | Reverse logic                                 |
| `LET()`        | Advanced    | Create variables for clean formulas           |

---

## 🔍 INTRODUCTION TO LOOKUP FUNCTIONS

Lookup functions help you **find data in a table or range** based on a lookup value. Different functions work in different directions and have unique capabilities.

We’ll cover:

1. **XLOOKUP()** – modern, flexible lookup
2. **VLOOKUP()** – vertical lookup
3. **HLOOKUP()** – horizontal lookup
4. **MATCH()** – position of a value
5. **INDEX()** – value at a position
6. **Combining INDEX() and MATCH()** – powerful alternative to VLOOKUP

---

## ✅ 1. `XLOOKUP()` – The Modern Replacement

### ✅ Purpose:

Finds a value in a row or column and returns a corresponding value. Works **vertically and horizontally**, supports **exact match by default**.

### ✅ Syntax:

```excel
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

- `lookup_value`: What you're searching for
- `lookup_array`: Where to search
- `return_array`: Where to return the result from
- `if_not_found`: Optional message if not found
- `match_mode`: 0=exact match (default), -1=next smaller, 1=next larger
- `search_mode`: 1=first to last, -1=last to first

### ✅ Example:

```excel
=XLOOKUP("Banana", A2:A10, B2:B10, "Not Found")
```

Looks for "Banana" in `A2:A10`, returns corresponding value from `B2:B10`.

---

## ✅ 2. `VLOOKUP()` – Vertical Lookup

### ✅ Purpose:

Searches for a value in the **first column** of a table and returns a value from another column **in the same row**.

### ✅ Syntax:

```excel
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

- `col_index_num`: The column number in the table to return
- `range_lookup`: FALSE = exact match, TRUE = approximate match

### ✅ Example:

```excel
=VLOOKUP("Banana", A2:C10, 2, FALSE)
```

Looks in column A for "Banana" and returns the value from column B (2nd column in range).

> ⚠️ Limitation: VLOOKUP **only searches to the right**, and it’s slower than XLOOKUP.

---

## ✅ 3. `HLOOKUP()` – Horizontal Lookup

### ✅ Purpose:

Looks for a value in the **top row** of a table and returns a value from a specified row **in the same column**.

### ✅ Syntax:

```excel
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

### ✅ Example:

```excel
=HLOOKUP("Sales", A1:G3, 2, FALSE)
```

Searches for "Sales" in row 1, returns value from row 2 in that column.

> ⚠️ Like VLOOKUP, HLOOKUP has limited flexibility. Better replaced by `XLOOKUP()` or `INDEX()` + `MATCH()`.

---

## ✅ 4. `MATCH()` – Find Position

### ✅ Purpose:

Returns the **relative position** of an item in a range.

### ✅ Syntax:

```excel
MATCH(lookup_value, lookup_array, [match_type])
```

- `match_type`: 0 = exact match, 1 = less than, -1 = greater than

### ✅ Example:

```excel
=MATCH("Banana", A2:A10, 0)
```

Returns the position of "Banana" in the range.

---

## ✅ 5. `INDEX()` – Return Value at a Position

### ✅ Purpose:

Returns a value at a specified **row and column** in a range or array.

### ✅ Syntax:

```excel
INDEX(array, row_num, [column_num])
```

### ✅ Example:

```excel
=INDEX(B2:B10, 3)
```

Returns the 3rd value from range `B2:B10`.

Or in 2D:

```excel
=INDEX(A2:C5, 2, 3)
```

Returns value from **2nd row and 3rd column** of the range.

---

## ✅ 6. `INDEX()` + `MATCH()` – Powerful Combo

This combination replaces VLOOKUP and overcomes its limitations:

- Works **left or right**
- More flexible with large data
- More efficient in large datasets

### ✅ Example:

**Table:**

| A       | B     |
| ------- | ----- |
| Product | Price |
| Apple   | 100   |
| Banana  | 60    |
| Mango   | 90    |

```excel
=INDEX(B2:B4, MATCH("Banana", A2:A4, 0))
```

- `MATCH("Banana", A2:A4, 0)` returns 2
- `INDEX(B2:B4, 2)` returns 60

---

## 🔁 COMPARISON TABLE

| Function    | Direction             | Return Type | Notes                                  |
| ----------- | --------------------- | ----------- | -------------------------------------- |
| `XLOOKUP()` | Horizontal + Vertical | Value       | Most powerful, replaces older lookups  |
| `VLOOKUP()` | Vertical              | Value       | Only searches right                    |
| `HLOOKUP()` | Horizontal            | Value       | Only searches down                     |
| `MATCH()`   | NA                    | Position    | Use with `INDEX()`                     |
| `INDEX()`   | NA                    | Value       | Use with `MATCH()` for flexible lookup |

---

## 🧪 PRACTICE EXERCISES

**Given Table:**

| Product | Price | Stock |
| ------- | ----- | ----- |
| Apple   | 100   | 25    |
| Banana  | 60    | 40    |
| Mango   | 90    | 35    |

1. **XLOOKUP the stock of “Banana”:**

```excel
=XLOOKUP("Banana", A2:A4, C2:C4)
```

2. **VLOOKUP the price of “Mango”:**

```excel
=VLOOKUP("Mango", A2:C4, 2, FALSE)
```

3. **INDEX + MATCH to get price of “Apple”:**

```excel
=INDEX(B2:B4, MATCH("Apple", A2:A4, 0))
```

4. **Find the column position of "Stock":**

```excel
=MATCH("Stock", A1:C1, 0)
```

---

Let's discuss on how to work with **date and time functions in Excel**, focusing on:

1. `NOW()`
2. `TODAY()`
3. `WEEKDAY()`
4. `WORKDAY()`

These functions are essential for **referencing the current date/time** and **calculating dates based on working days and weekdays**.

---

## ✅ 1. `NOW()` Function

### ✅ Purpose:

Returns the **current date and time**.

### ✅ Syntax:

```excel
=NOW()
```

### ✅ Example:

If today is August 7, 2025, and the time is 10:45 AM, then:

```excel
=NOW()
```

might return:

```
07-08-2025 10:45 AM
```

### ✅ Notes:

- It **updates automatically** whenever the worksheet recalculates.
- Format the cell as **Date/Time** or **Custom** (`dd-mm-yyyy hh:mm AM/PM`) to show both parts.

---

## ✅ 2. `TODAY()` Function

### ✅ Purpose:

Returns the **current date** (no time).

### ✅ Syntax:

```excel
=TODAY()
```

### ✅ Example:

```excel
=TODAY()
```

might return:

```
07-08-2025
```

### ✅ Notes:

- Like `NOW()`, it auto-updates every day.
- Useful for calculating due dates, age, aging of invoices, etc.

---

## 🧮 Example: Calculate Age

```excel
=DATEDIF(B2, TODAY(), "Y")
```

Returns the age (in years) based on birth date in B2.

---

## ✅ 3. `WEEKDAY()` Function

### ✅ Purpose:

Returns a number representing the **day of the week** for a given date.

### ✅ Syntax:

```excel
WEEKDAY(serial_number, [return_type])
```

- `serial_number`: The date you want to evaluate.
- `return_type` (optional): Determines what number corresponds to which day.

### ✅ Return Types:

| Return Type | Week Starts On | Sunday = | Monday = |
| ----------- | -------------- | -------- | -------- |
| 1 (default) | Sunday         | 1        | 2        |
| 2           | Monday         | 7        | 1        |
| 3           | Monday         | 0        | 1        |

### ✅ Example:

```excel
=WEEKDAY(TODAY())
```

If today is Thursday, it returns:

```
5  (if using default return_type = 1)
```

### ✅ Example with Custom Return Type:

```excel
=WEEKDAY(A1, 2)
```

- If A1 = 07-08-2025 (Thursday), result = 4 (Monday = 1)

---

### ✅ Use Case: Label Day of the Week

```excel
=CHOOSE(WEEKDAY(A1, 2), "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
```

---

## ✅ 4. `WORKDAY()` Function

### ✅ Purpose:

Returns a date that is a specified number of **working days (excluding weekends and holidays)** before or after a given start date.

### ✅ Syntax:

```excel
WORKDAY(start_date, days, [holidays])
```

- `start_date`: Starting date
- `days`: Number of workdays to add (can be negative)
- `holidays`: Optional range of holiday dates to exclude

### ✅ Example 1: Add 10 working days

```excel
=WORKDAY(TODAY(), 10)
```

If today is 07-Aug-2025, result might be:

```
21-Aug-2025
```

### ✅ Example 2: Add workdays with holidays

```excel
=WORKDAY(A1, 5, B1:B3)
```

Where:

- `A1` = Start date
- `B1:B3` = List of holidays

---

## 🔁 COMPARISON OF FUNCTIONS

| Function    | Description                                                       | Returns      | Auto-updates |
| ----------- | ----------------------------------------------------------------- | ------------ | ------------ |
| `NOW()`     | Current date + time                                               | DateTime     | Yes          |
| `TODAY()`   | Current date                                                      | Date only    | Yes          |
| `WEEKDAY()` | Day of week number for a date                                     | Number (1–7) | No           |
| `WORKDAY()` | Future/past workday after skipping weekends and optional holidays | Date         | No           |

---

## 🧪 PRACTICE EXERCISES

### 1. **Calculate Due Date 10 Working Days After Invoice Date**

```excel
=WORKDAY(A2, 10)
```

### 2. **Get Today’s Day of the Week (e.g., “Thursday”)**

```excel
=TEXT(TODAY(), "dddd")
```

### 3. **Check if a Date is on a Weekend**

```excel
=IF(WEEKDAY(A2, 2)>5, "Weekend", "Weekday")
```

### 4. **Find Date 20 Workdays Before a Given Date**

```excel
=WORKDAY(A2, -20)
```

---

## 🔧 TIPS

- Use `NOW()` and `TODAY()` in dashboards and templates where you want the current time or date.
- Use `WORKDAY()` for **project deadlines**, **invoice due dates**, or **leave management**.
- Combine `WEEKDAY()` with `IF()` or `CHOOSE()` to control logic based on day type.

---

## 📌 WHAT IS CONSOLIDATE?

The **Consolidate** feature in Excel allows you to **combine and summarize data from different worksheets or ranges**, using functions like **SUM, AVERAGE, COUNT, MAX, MIN**, etc.

It's particularly useful when:

- You have the **same structure** of data across multiple sheets (e.g., monthly reports).
- You want a **summary table** from those multiple ranges.

---

## ✅ SCENARIOS WHERE CONSOLIDATE IS USEFUL

| Sheet Name | Data Type         | Example                 |
| ---------- | ----------------- | ----------------------- |
| Sheet1     | January Sales     | Region-wise sales       |
| Sheet2     | February Sales    | Same format             |
| Sheet3     | March Sales       |                         |
| Summary    | Consolidated data | Total quarterly figures |

---

## ✅ HOW TO USE THE CONSOLIDATE FEATURE

### 🔹 STEP 1: Prepare Your Data

- Ensure **data ranges have identical structure**: Same row and column labels.
- Example:

**Sheet1 (January)**

| Region | Sales |
| ------ | ----- |
| East   | 100   |
| West   | 200   |

**Sheet2 (February)**

| Region | Sales |
| ------ | ----- |
| East   | 150   |
| West   | 180   |

---

### 🔹 STEP 2: Create a Summary Sheet

- Insert a new sheet (e.g., named `Summary`).
- Click in the top-left cell where you want the consolidated data to appear.

---

### 🔹 STEP 3: Open the Consolidate Dialog

1. Go to the **Data** tab on the Ribbon.
2. Click on **Consolidate** (in the **Data Tools** group).

---

### 🔹 STEP 4: Configure Consolidation

The **Consolidate dialog box** appears.

#### Choose a Function:

- Select a function to apply to the data:

  - `SUM` (default)
  - `AVERAGE`
  - `COUNT`
  - `MAX`, `MIN`
  - `PRODUCT`, etc.

#### Add References:

1. Click in the **Reference** box.
2. Go to the first sheet (e.g., `Sheet1`) and select the data range (e.g., `A1:B3`).
3. Click **Add**.
4. Repeat for `Sheet2`, `Sheet3`, etc.

#### Use Labels:

- Check:

  - **Top row** → If column labels are in the first row.
  - **Left column** → If row labels (like Region names) are in the first column.

---

### 🔹 STEP 5: Click OK

- Excel will generate a **summary table** in the selected location.
- It will **combine the data by matching labels** and apply the selected function.

---

## 🔁 EXAMPLE

| Region | Jan | Feb |
| ------ | --- | --- |
| East   | 100 | 150 |
| West   | 200 | 180 |

If you consolidate using `SUM`, the summary will be:

| Region | Total |
| ------ | ----- |
| East   | 250   |
| West   | 380   |

---

## ✅ TIPS AND TRICKS

| Tip                               | Description                                                |
| --------------------------------- | ---------------------------------------------------------- |
| Keep data structure consistent    | Same rows and column labels across sheets                  |
| Use named ranges                  | Makes it easier to add ranges in the dialog                |
| Use dynamic ranges                | With tables or dynamic named ranges, updates become easier |
| Use “Create links to source data” | Adds links and groupings to trace back to original data    |

---

## 🔧 LIMITATIONS

- Works best when data layout is **identical**.
- Not dynamic: **changes in source data won’t auto-update** consolidated result unless links are created.
- Cannot summarize data across **non-contiguous structures** easily.

---

## 🧪 PRACTICE TASK

1. Create three sheets: **Q1, Q2, Q3** with same headers: Product | Sales
2. Fill each with different numbers.
3. On a new sheet, use **Data → Consolidate → SUM** to combine sales by product.
4. Try again using **AVERAGE**.

---

## 🎯 WHAT IS “WHAT-IF ANALYSIS”?

**What-If Analysis** in Excel allows you to **experiment with different sets of input values** and see how they affect the result of formulas on your worksheet.

One of the most powerful tools for this is the **Scenario Manager**, which lets you define and compare **multiple scenarios** with different variable values — like Best Case, Worst Case, or Custom Cases.

---

## ✅ WHAT IS SCENARIO MANAGER?

**Scenario Manager** lets you:

- Define **multiple groups of input values** (called _scenarios_)
- Instantly **switch between these scenarios** to view results
- Summarize all scenarios in a **comparison table**

---

## ✅ WHEN TO USE SCENARIO MANAGER?

Use Scenario Manager when:

- You want to **analyze several business cases** (e.g., profit projections)
- You want to **save different combinations of values** for quick reference
- You want to **compare outcomes side-by-side** based on different inputs

---

## 🔧 EXAMPLE SCENARIO

You run a business and want to forecast **Net Profit**, which depends on:

- **Selling Price**
- **Cost Price**
- **Units Sold**

Your formula:

```excel
Net Profit = (Selling Price - Cost Price) * Units Sold
```

---

## 🧭 STEP-BY-STEP: USING SCENARIO MANAGER

---

### 🔹 STEP 1: Prepare the Worksheet

Create a sheet like this:

| Cell | Label         | Value        |
| ---- | ------------- | ------------ |
| B1   | Selling Price | 500          |
| B2   | Cost Price    | 300          |
| B3   | Units Sold    | 100          |
| B4   | Profit        | =(B1-B2)\*B3 |

So, Profit is dynamically calculated based on inputs.

---

### 🔹 STEP 2: Open Scenario Manager

1. Go to the **Data** tab.
2. Click **What-If Analysis** (in the _Forecast_ group).
3. Choose **Scenario Manager**.

---

### 🔹 STEP 3: Add a Scenario

1. Click **Add**.
2. Name the scenario — e.g., **Best Case**.
3. In **Changing cells**, select B1, B2, B3 (input cells).
4. Click **OK**.
5. Enter values for this scenario, e.g.:

   - Selling Price: 600
   - Cost Price: 250
   - Units Sold: 150

6. Click **OK** or **Add** another scenario.

Repeat for:

- **Worst Case**: e.g., 450 / 320 / 80
- **Expected Case**: e.g., 500 / 300 / 100

---

### 🔹 STEP 4: Show a Scenario

In Scenario Manager:

- Select a scenario name.
- Click **Show**.
- Excel updates the sheet values to reflect the chosen scenario, recalculating the result.

---

### 🔹 STEP 5: Create a Scenario Summary

1. In Scenario Manager, click **Summary**.
2. Choose:

   - **Scenario Summary**
   - Result cells: select `B4` (Profit cell)

3. Click **OK**.

📋 Excel creates a **new sheet** summarizing each scenario and its effect on the result.

---

## 🧪 EXAMPLE OUTPUT

| Changing Cells | Selling Price | Cost Price | Units Sold | Profit |
| -------------- | ------------- | ---------- | ---------- | ------ |
| Best Case      | 600           | 250        | 150        | 52500  |
| Worst Case     | 450           | 320        | 80         | 10400  |
| Expected Case  | 500           | 300        | 100        | 20000  |

---

## ✅ ADVANTAGES OF USING SCENARIO MANAGER

| Feature                    | Benefit                                |
| -------------------------- | -------------------------------------- |
| Stores multiple cases      | Easily switch and review outcomes      |
| Editable anytime           | Update scenarios as assumptions change |
| Auto summary               | Compare results in a separate sheet    |
| Useful for decision-making | Visualize best/worst case projections  |

---

## 🧠 TIPS

- Use **descriptive names** for each scenario (e.g., "High Demand", "Low Pricing").
- Keep your **input cells grouped** logically for easy selection.
- Combine with **Data Validation** to limit inputs to valid ranges.

---

## 🚫 LIMITATIONS

- Scenario Manager works best for **up to 32 changing cells** per scenario.
- Doesn’t update **automatically** when inputs change — you must Show or Refresh scenarios.

---

## ✅ SUMMARY TABLE

| Feature      | Description                            |
| ------------ | -------------------------------------- |
| What-It Does | Tests different combinations of inputs |
| Max Inputs   | Up to 32 cells per scenario            |
| Output       | Summary table of inputs and results    |
| Used For     | Forecasting, budgeting, planning       |

---

## ✅ **1. Calculate Financial Data Using the `PMT()` Function**

### **Purpose**:

The `PMT()` function calculates the payment for a loan based on constant payments and a constant interest rate.

### **Syntax**:

```excel
PMT(rate, nper, pv, [fv], [type])
```

### **Parameters**:

- `rate` – Interest rate for each period (monthly rate if monthly payments).
- `nper` – Total number of payment periods.
- `pv` – Present value, or the total amount of the loan.
- `fv` _(optional)_ – Future value, or the cash balance you want to attain after the last payment. Defaults to `0`.
- `type` _(optional)_ – When payments are due:

  - `0` – End of the period (default)
  - `1` – Beginning of the period

### **Example**:

You borrow ₹5,00,000 at an annual interest rate of 9% for 5 years and want to find the monthly payment.

```excel
=PMT(9%/12, 5*12, -500000)
```

- Interest rate per month: `9%/12`
- Number of months: `5*12`
- Loan amount: `₹500,000` (negative because it’s an outgoing payment)

🟢 **Result**: Approx. `₹10,377.44` (monthly payment)

---

## ✅ **2. Filter Data Using the `FILTER()` Function**

### **Purpose**:

`FILTER()` returns an array of values that meet specified criteria.

### **Syntax**:

```excel
FILTER(array, include, [if_empty])
```

### **Parameters**:

- `array` – The range of data to filter.
- `include` – The condition(s) that determine which data to return.
- `if_empty` _(optional)_ – What to return if no data matches.

### **Example**:

Filter sales where the amount is greater than ₹10,000.

Assume:

- A2\:A10 – Product names
- B2\:B10 – Sales amount

```excel
=FILTER(A2:B10, B2:B10>10000, "No Match")
```

🟢 This will return rows where sales are above ₹10,000.

---

## ✅ **3. Sort Data Using the `SORTBY()` Function**

### **Purpose**:

`SORTBY()` allows you to sort data based on the values in one or more columns.

### **Syntax**:

```excel
SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)
```

### **Parameters**:

- `array` – Range of data to sort.
- `by_array1` – The first column/range to sort by.
- `sort_order1` – `1` for ascending, `-1` for descending.

### **Example**:

Sort sales data by amount in descending order.

Assume:

- A2\:A10 – Product names
- B2\:B10 – Sales amount

```excel
=SORTBY(A2:B10, B2:B10, -1)
```

🟢 This will display the product and sales data, sorted by sales in descending order.

---

## 🔁 **Combined Example**:

Let’s combine them into a mini-report:

- You want to list all products with sales above ₹10,000.
- Then sort them in descending order of sales.

```excel
=SORTBY(FILTER(A2:B10, B2:B10>10000), B2:B10, -1)
```

---

## ✅ Summary Table:

| Function   | Use Case                                   | Example Syntax                  |
| ---------- | ------------------------------------------ | ------------------------------- |
| `PMT()`    | Calculate EMI or loan payments             | `=PMT(9%/12, 60, -500000)`      |
| `FILTER()` | Return only rows matching criteria         | `=FILTER(A2:B10, B2:B10>10000)` |
| `SORTBY()` | Sort filtered or full data based on column | `=SORTBY(A2:B10, B2:B10, -1)`   |

---

Creating and modifying **advanced charts** in **Microsoft Excel** allows users to visualize complex data trends, comparisons, and patterns more effectively than with basic charts. Below is a detailed explanation of how to create and customize advanced charts in Excel:

---

## ✅ **1. What Are Advanced Charts?**

Advanced charts go beyond basic types like column, bar, or pie charts. They include:

- Combo charts (dual-axis)
- Histogram
- Pareto chart
- Waterfall chart
- Funnel chart
- Radar chart
- Bubble chart
- Stock chart
- Surface chart
- Treemap and Sunburst

These charts help present multidimensional or hierarchical data and provide deep insights.

---

## ✅ **2. Creating Advanced Charts**

### ✦ A. **Combo Chart (Custom Combination of Charts)**

**Use Case:** Compare two data series with different units (e.g., Revenue in ₹ vs. Growth in %)

**Steps:**

1. Select your data range (e.g., Product | Revenue | Growth Rate).
2. Go to **Insert > Combo Chart**.
3. Choose “Custom Combination Chart.”
4. Set one series to **Column**, another to **Line** (or any combo).
5. Check “Secondary Axis” for the appropriate series (e.g., Growth Rate).
6. Click **OK**.

---

### ✦ B. **Histogram**

**Use Case:** Analyze data distribution (e.g., exam scores).

**Steps:**

1. Select the numeric data.
2. Go to **Insert > Insert Statistic Chart > Histogram**.
3. Customize bin width (Right-click horizontal axis > Format Axis > Bin Width).

---

### ✦ C. **Pareto Chart**

**Use Case:** 80/20 analysis (e.g., defects vs. frequency).

**Steps:**

1. Create a table with **Categories** and **Frequency**.
2. Sort data in descending order of frequency.
3. Insert a **Histogram > Pareto Chart**.
4. Excel automatically adds a cumulative percentage line.

---

### ✦ D. **Waterfall Chart**

**Use Case:** Show how values add up over time (e.g., Profit calculation).

**Steps:**

1. Create a data table with headings like “Start”, “Revenue”, “Cost”, “Profit.”
2. Go to **Insert > Waterfall or Stock Chart > Waterfall**.
3. Click data bars and mark totals manually using “Set as Total.”

---

### ✦ E. **Funnel Chart**

**Use Case:** Track stages of a sales pipeline or conversion process.

**Steps:**

1. Data format: Stage | Value (e.g., Leads → Qualified → Proposal → Closed).
2. Insert > Funnel Chart.

---

### ✦ F. **Radar Chart**

**Use Case:** Compare multivariable data for performance evaluation.

**Steps:**

1. Arrange data: Categories as rows/columns (e.g., Skills), values in other cells.
2. Go to **Insert > Radar Chart** (with or without markers).

---

### ✦ G. **Bubble Chart**

**Use Case:** Show 3 dimensions of data (X, Y, and Size).

**Steps:**

1. Data format: X values | Y values | Bubble size.
2. Insert > Scatter Chart > Bubble Chart.

---

### ✦ H. **Treemap & Sunburst**

**Use Case:** Visualize hierarchical data (like categories and subcategories).

**Steps:**

1. Format: Category | Subcategory | Value.
2. Insert > Treemap or Sunburst chart.

---

## ✅ **3. Modifying Advanced Charts**

### ✦ A. **Change Chart Type**

- Select the chart.
- Go to **Chart Design > Change Chart Type**.

### ✦ B. **Add or Modify Chart Elements**

Go to **Chart Design > Add Chart Element**, and choose from:

- Axis Titles
- Data Labels
- Legends
- Gridlines
- Trendlines

---

### ✦ C. **Format Chart Elements**

- Right-click on chart elements (axes, bars, lines, etc.).
- Choose **Format \[Element]** (e.g., Format Axis, Format Data Series).
- Customize:

  - Color, Fill, Outline
  - Shadow, Glow, Transparency
  - Data label position and content
  - Axis bounds, intervals, and number formatting

---

### ✦ D. **Use Custom Number Formats**

For example:

- Show percentages with no decimal: `0%`
- Currency with symbol: `₹#,##0`

---

### ✦ E. **Create a Secondary Axis**

Useful in combo charts:

- Click data series > Format Data Series > Plot Series on Secondary Axis.

---

### ✦ F. **Apply Conditional Formatting (In-Chart)**

Not natively supported in charts, but achievable by:

- Pre-formatting the source data conditionally.
- Using formulas to create separate series for color-coding.

---

### ✦ G. **Apply Themes and Styles**

- Use **Chart Design > Chart Styles** to apply predefined themes.
- Or customize colors using **Format > Shape Fill / Outline**.

---

### ✦ H. **Insert Interactive Elements**

- Use **Slicers**, **Timeline filters**, or **Dropdowns** from **Developer > Form Controls** to make charts dynamic (with help of pivot tables or formulas).

---

## ✅ **4. Tips for Better Charts**

- **Use labels and legends effectively** to make the chart readable.
- **Avoid clutter** – don’t use too many data series in one chart.
- **Use color consistently** to maintain clarity.
- **Add data callouts or annotations** to highlight important points.

---

## ✅ Summary Table

| Chart Type | Best For               | Key Feature         |
| ---------- | ---------------------- | ------------------- |
| Combo      | Multiple data types    | Dual axes           |
| Histogram  | Distribution analysis  | Bins                |
| Pareto     | 80/20 rule             | Cumulative line     |
| Waterfall  | Running totals         | Intermediate values |
| Funnel     | Staged processes       | Top-down view       |
| Radar      | Performance comparison | Spokes              |
| Bubble     | 3D data                | Size-based          |
| Treemap    | Hierarchical data      | Nested boxes        |

---

## ✅ **Create and Modify Dual Axis Charts**

### What is a Dual Axis Chart?

A dual-axis chart shows two different data series on the same chart with two different vertical axes (Y-axes) – left and right – making it easier to compare datasets with different value ranges.

### When to Use:

- Compare values with different units/scales (e.g., revenue vs. units sold).
- Show correlation between two data series.

### How to Create a Dual Axis Chart:

1. **Prepare your data** – e.g.:

   | Month | Revenue | Units Sold |
   | ----- | ------- | ---------- |
   | Jan   | 50,000  | 300        |
   | Feb   | 55,000  | 280        |

2. **Insert a Chart**:

   - Select the data.
   - Go to `Insert` > `Combo Chart` > `Custom Combo Chart`.
   - Set:

     - Revenue as **Clustered Column**
     - Units Sold as **Line**
     - Check the box: **Secondary Axis** for Units Sold.

3. **Customize**:

   - Change colors, labels, and add data labels as needed.
   - You can adjust each axis independently under **Chart Tools > Format Axis**.

---

## 📊 **Create and Modify Specialized Charts**

### 1. **Box & Whisker Chart**

- Used to show distribution of data (min, 1st quartile, median, 3rd quartile, max).

#### Steps:

1. Select your dataset (column-wise data).
2. Go to `Insert` > `Insert Statistic Chart` > `Box and Whisker`.
3. Excel will automatically calculate and draw the box plot.

---

### 2. **Combo Chart**

- Combines two chart types like column + line (used in Dual Axis Charts too).

#### Steps:

1. Select your data.
2. Go to `Insert` > `Combo Chart` > choose types for each series.
3. Choose which series to place on the secondary axis.

---

### 3. **Funnel Chart**

- Represents stages in a process, like sales pipelines or conversions.

#### Steps:

1. Use descending order data:

   | Stage     | Count |
   | --------- | ----- |
   | Prospects | 1000  |
   | Contacted | 600   |
   | Converted | 250   |

2. Select data > Insert > Funnel Chart.

---

### 4. **Histogram**

- Shows frequency distribution of data (e.g., number of students by score range).

#### Steps:

1. Select numeric data.
2. Go to `Insert` > `Insert Statistic Chart` > `Histogram`.

#### Customize:

- Right-click X-axis > **Format Axis** to set bin width or number of bins.

---

### 5. **Sunburst Chart**

- Hierarchical chart; shows levels of data (e.g., category → subcategory).

#### Example:

| Category | Subcategory | Value |
| -------- | ----------- | ----- |
| Tech     | Phones      | 50    |
| Tech     | Laptops     | 30    |
| Food     | Fruits      | 40    |

- Create hierarchy using multiple columns.
- Select data > Insert > Sunburst.

---

### 6. **Waterfall Chart**

- Used for tracking cumulative values (profit/loss, budget changes).

#### Steps:

1. Use a table like:

   | Category | Amount |
   | -------- | ------ |
   | Start    | 10000  |
   | Income   | 5000   |
   | Expenses | -3000  |
   | Final    | 12000  |

2. Select data > Insert > Waterfall Chart.

3. Set “Total” values manually by right-clicking on bars > Set as Total.

---

## 🔹 **Sparklines in Excel**

### What are Sparklines?

Tiny charts inside a single cell, used to give a visual trend of data (like mini line graphs).

### Types of Sparklines:

- Line
- Column
- Win/Loss

---

## ✅ **Inserting Sparklines**

### Steps:

1. Select data (e.g., sales over months).
2. Go to `Insert` > `Sparklines` (choose Line/Column/Win-Loss).
3. Choose location range (cells where sparklines appear).
4. Click OK.

---

## 🎨 **Customizing Sparklines**

- **Design Tab** appears when you select a cell with a sparkline.
- Customize with:

  - **Marker Options** – highlight high, low, first, last points.
  - **Sparkline Color** – change line or column color.
  - **Axis Options** – fix min/max axis values.
  - **Group** – edit multiple sparklines together or separately.

---

## Summary Table

| Chart Type    | Purpose                              | Key Feature                        |
| ------------- | ------------------------------------ | ---------------------------------- |
| Dual Axis     | Compare different scales of data     | Two Y-axes                         |
| Box & Whisker | Show data distribution               | Quartiles, median, outliers        |
| Combo         | Mix chart types for comparison       | E.g., column + line                |
| Funnel        | Show progressive reduction in stages | Best for sales pipelines           |
| Histogram     | Frequency distribution               | Auto bin grouping                  |
| Sunburst      | Hierarchical, circular view          | Category-subcategory relationships |
| Waterfall     | Running total across values          | Positive/negative changes          |
| Sparklines    | Tiny in-cell visualizations          | Great for dashboards               |

---

## **What is a Pivot Table?**

A **Pivot Table** is an interactive tool in Excel that allows you to **summarize, analyze, explore, and present** large amounts of data. It helps in extracting meaningful insights by performing operations like grouping, sorting, counting, averaging, or displaying data trends.

---

## **1. Creating Pivot Tables**

### Steps:

1. **Select your data**: The dataset should have headers in the first row.
2. Go to the **Insert tab** → Click on **PivotTable**.
3. In the dialog box:

   - Choose the data range.
   - Choose whether to place the Pivot Table in a **new worksheet** or **existing worksheet**.

4. Click **OK**.

---

## **2. PivotTable – Fields**

After creating a Pivot Table, the **PivotTable Field List** appears on the right side:

- Lists all the **column headers** from your data source.
- You can **drag and drop** these fields into one of the four areas:

  - **Filters**
  - **Columns**
  - **Rows**
  - **Values**

---

## **3. PivotTable Layout - Fields and Areas**

### The four main areas:

- **Filters**: Used to filter the entire PivotTable based on a specific field (e.g., Region, Year).
- **Columns**: Field values appear as **column headers**.
- **Rows**: Field values appear as **row headers**.
- **Values**: Contains the data to be **summarized** (Sum, Count, Average, etc.)

> Example: If you place `Region` in **Rows**, `Month` in **Columns**, and `Sales` in **Values**, the Pivot Table shows total sales per region per month.

---

## **4. Nesting, Expanding, and Collapsing Fields**

### Nesting:

- You can **nest fields** by dragging more than one field to Rows or Columns.
  E.g., `Region` and then `City` in Rows to view sales per city within each region.

### Expand/Collapse:

- Click the **+/- buttons** next to a field to **expand or collapse** grouped or nested data.
- Right-click → **Expand/Collapse** to control levels.

---

## **5. Grouping and Ungrouping Field Values**

You can **group**:

- **Dates** (by days, months, quarters, years)
- **Numbers** (by ranges like 0–100, 101–200, etc.)
- **Text** (manual grouping)

### How:

1. Select a field value in PivotTable.
2. Right-click → **Group**.
3. Choose grouping options.

To **Ungroup**:

- Right-click grouped item → **Ungroup**.

---

## **6. PivotTable – Reports**

You can **summarize data** using different report types:

- **Sum**, **Count**, **Average**, **Max**, **Min**, etc.
- Change calculation by:

  - Right-click a value in Pivot → **Summarize Values By**
  - Or use **Value Field Settings**

You can also **show values as**:

- % of Total
- % Difference From
- Running Total, Rank, etc.

---

## **7. Inserting Slicers**

**Slicers** are visual filters that make it easier to filter Pivot Tables.

### Steps:

1. Click on the Pivot Table.
2. Go to **PivotTable Analyze** tab → Click **Insert Slicer**.
3. Select the field you want to filter by (e.g., Region).
4. Click OK → Slicer appears on the sheet.

---

## **8. Multi-Select Option in Slicers**

By default, you can **select one item** in a slicer.
To **select multiple items**:

- Click the **multi-select icon** (a checkbox icon) on the slicer.
- Hold **Ctrl** (or click multiple options if multi-select is enabled).

---

## **9. PivotTable Enhancements**

Advanced functionalities include:

- **Refreshing data** when the source data changes.
- **Changing source data range** via PivotTable Options → Change Data Source.
- **Show Report Filter Pages**: Automatically create separate sheets based on a filter.
- **Calculated Fields**: Add formulas to a Pivot Table.

---

## **10. Inserting Pivot Charts**

Pivot Charts are graphical representations of Pivot Table data.

### Steps:

1. Select Pivot Table.
2. Go to **Insert tab** → Click **Pivot Chart**.
3. Choose chart type (Column, Bar, Line, etc.)
4. Click OK.

The chart is **linked to the Pivot Table** – changing the table changes the chart.

---

## **11. More Pivot Table Functionality**

- **Calculated Fields**: Use your own formulas in PivotTables.

  - PivotTable Analyze → Fields, Items & Sets → Calculated Field

- **Conditional Formatting**: Apply color scales or rules to values in Pivot.
- **PivotTable Styles**: Format tables using predefined or custom styles.
- **Field Settings**: Customize how values are summarized or labeled.

---

## **12. Working with Pivot Tables – Best Practices**

- Use **tables** (Ctrl + T) as source data to auto-expand with new data.
- Always **refresh** after updating source data (Alt + F5).
- Use **slicers** or **timelines** for easy interactivity.
- Keep data **clean and consistent** – no merged cells, blanks in headers, or inconsistent data types.
- Use **dynamic named ranges** or **Excel Tables** for scalable PivotTables.

---

### 🔍 **What is Worksheet Auditing in Excel?**

Auditing in Excel helps you **trace and understand the relationships between cells** – especially useful in large or complex spreadsheets where formulas pull data from multiple locations. It visually shows where a formula gets its inputs or where its results are used.

---

## 1. ✅ **Tracing Precedents**

### 📌 **Definition:**

Precedents are **cells that provide data to a formula** in the active cell. Tracing precedents shows **arrows pointing to the cells** that the current formula cell depends on.

### 🧩 **Why Use It:**

- To see **where a formula pulls data from**
- To troubleshoot **incorrect results**
- To analyze dependencies before modifying formulas

### ▶️ **How to Use:**

1. **Select the cell** with the formula.
2. Go to the **Formulas** tab on the Ribbon.
3. In the **Formula Auditing** group, click **“Trace Precedents.”**
4. Blue arrows appear pointing to the precedent cells.

   - **Solid line:** data from same sheet.
   - **Dashed line:** data from another worksheet or workbook.

### 🧹 To remove arrows:

- Click **“Remove Arrows”** under the Formula Auditing group.

---

## 2. 🔄 **Tracing Dependents**

### 📌 **Definition:**

Dependents are **cells that use the current cell’s value in their formulas**. Tracing dependents shows **arrows from the current cell to other cells** that rely on it.

### 🧩 **Why Use It:**

- To identify **where changes in a cell might affect results**
- To prevent **breaking formulas** in other areas
- To help with **impact analysis**

### ▶️ **How to Use:**

1. Select the cell.
2. Go to **Formulas** → **Trace Dependents**.
3. Arrows point to cells that are **affected** by the selected cell.

---

## 3. 👀 **Showing Formulas**

### 📌 **Definition:**

Instead of showing results, this feature displays **all the actual formulas** in cells. It helps identify what formulas are used and where.

### ▶️ **How to Use:**

- Go to the **Formulas** tab.
- Click on **“Show Formulas”** in the Formula Auditing group.
- All formulas in the worksheet will be **displayed instead of results**.
- Click again to **toggle back to normal view**.

### 🧩 **Alternate Shortcut:**

Press **`Ctrl` + `~` (tilde)** to toggle Show Formulas on/off.

---

## 🧠 Bonus Tips:

- Use **Evaluate Formula** in the same Formula Auditing group to **step through formula calculations.**
- Use **Watch Window** to monitor important cells across large spreadsheets.

---

### 📊 Summary Table:

| Feature          | Purpose                                  | Access Path                 | Shortcut (if any) |
| ---------------- | ---------------------------------------- | --------------------------- | ----------------- |
| Trace Precedents | See source cells for a formula           | Formulas → Trace Precedents | —                 |
| Trace Dependents | See which cells rely on the current cell | Formulas → Trace Dependents | —                 |
| Show Formulas    | View formulas instead of results         | Formulas → Show Formulas    | `Ctrl` + `~`      |
| Remove Arrows    | Clear auditing arrows                    | Formulas → Remove Arrows    | —                 |

---
