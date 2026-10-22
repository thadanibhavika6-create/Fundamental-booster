# Fundamental-booster

<div align="center">

<img src="assets/header.svg" alt="Excel Data Analysis & Formulas Project" width="100%"/>

<img src="assets/typing.svg" alt="Typing animation" width="80%"/>

![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Sheets](https://img.shields.io/badge/Worksheets-3-blue?style=for-the-badge)
![Records](https://img.shields.io/badge/Records-60-orange?style=for-the-badge)
![Formulas](https://img.shields.io/badge/Functions-20%2B-purple?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=for-the-badge)

</div>

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 📌 Overview

This project is a hands-on Excel workbook (`pr1.xlsx`) that applies **logical, lookup, statistical, text and date functions** to three practical datasets. Each sheet solves business-style questions such as *"Who are the top students?"*, *"What are the total sales for a region?"* and *"What is an employee's salary?"* using clean, reusable formulas.

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🗂️ Workbook Structure

| # | Sheet | Records | What it covers |
|---|-------|:-------:|----------------|
| 1 | 🎓 **students grade** | 20 | Marks, totals, percentage, grades, top-student flag |
| 2 | 🛒 **sale data** | 20 | Products, regions, discounts, eligibility, sales lookups |
| 3 | 👥 **Employee data** | 20 | Departments, salaries, tenure, text & dynamic lookups |

All data is stored in **Excel Tables** (`Table1`, `Table4`, `Table5`) so formulas stay readable and expand automatically.

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🎓 Sheet 1 — Students Grade

**Columns:** Student ID · Name · Math · Science · English · Enrollment Date · Grade · Date of Birth · Score · Total Marks · Percentage

| Feature | Formula used |
|---------|--------------|
| Relative reference | `=C2` |
| Absolute reference | `=$C$2` |
| Grade (A–F) from average | `=IF(AVERAGE(C2:E2)>=90,"A",IF(AVERAGE(C2:E2)>=80,"B",...))` |
| Top Student flag | `=IF(AND(C2>80,D2>80),"Top Student","Normal")` |
| Total marks | `=C2+D2+E2` |
| Percentage | `=(C2+D2+E2)/300*100` |

**🧮 Analysis tasks solved**

| Task | Function | Result |
|------|----------|:------:|
| Students scoring above 50 in Math | `COUNTIF` | **19** |
| Average Math score of students above 60 | `AVERAGEIFS` | **82.26** |
| Fetch student name from ID | `VLOOKUP` | **Student 6** |
| List of top performers | Filtered extract | **10 students** |

**🏅 Grading scale**

| Average | Grade |
|:-------:|:-----:|
| 90 and above | 🥇 A |
| 80 – 89 | 🥈 B |
| 70 – 79 | 🥉 C |
| 60 – 69 | D |
| Below 60 | F |

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🛒 Sheet 2 — Sales Data

**Columns:** Sales ID · Product Code · Product · Region · Salesperson · Amount · Date · Discount Rate · Discount Type · Month

**💸 Discount slabs**

| Amount | Discount |
|:------:|:--------:|
| ≥ 40,000 | 20% |
| ≥ 20,000 | 10% |
| ≥ 10,000 | 5% |
| Below 10,000 | No Discount |

A sale is marked **Eligible** when the amount is **7,500 or more**.

**🧮 Analysis tasks solved**

| Task | Function | Result |
|------|----------|:------:|
| Total sales for a region + product (Mouse, East) | `SUMIFS` | **37,335** |
| Sales value for a product code (PC110) | `VLOOKUP` | **35,185** |
| Sales of a salesperson in a given month | `INDEX` + `MATCH` | **5,264** |
| Find salesperson by sales amount | `XLOOKUP` | **Person 12** |
| Position of a product in the list | `XMATCH` | **3** |

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 👥 Sheet 3 — Employee Data

**Columns:** Employee ID · Name · Department · Salary · Joining Date · Working Days

| Task | Function / Formula |
|------|--------------------|
| Working days since joining | `=DATEDIF(E2,TODAY(),"d")` |
| Extract first characters of a name | `=LEFT(B7,5)` |
| Find position of a character | `=FIND("m",B8)` |
| Convert to uppercase | `=UPPER(B5)` |
| Convert to lowercase | `=LOWER(B12)` |
| Two-way dynamic lookup | `=INDEX(A2:E21, MATCH("Employee 10",B2:B21,0), MATCH("Salary",A1:E1,0))` |
| Lookup on unsorted data | `=XLOOKUP("Employee 6",B2:B21,D2:D21)` |
| Dynamic range reference | `=SUM(INDIRECT("E2:E21"))` |

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🧰 Excel Skills Demonstrated

| Category | Functions |
|----------|-----------|
| 🔀 **Logical** | `IF` · `AND` · `OR` (nested) |
| 🔍 **Lookup** | `VLOOKUP` · `XLOOKUP` · `INDEX` · `MATCH` · `XMATCH` |
| 📈 **Statistical** | `COUNTIF` · `SUMIFS` · `AVERAGEIFS` · `AVERAGE` |
| 🔤 **Text** | `LEFT` · `FIND` · `UPPER` · `LOWER` · `TEXT` |
| 📅 **Date & Time** | `TODAY` · `DATEDIF` · `MONTH` |
| 🧩 **Advanced** | `INDIRECT` · Structured Table references · Absolute / relative references |

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🚀 How to Use

1. **Download** or clone this repository.
2. **Open** `pr1.xlsx` in Microsoft Excel (2021 / Microsoft 365 recommended).
3. Switch between the three sheets using the tabs at the bottom.
4. Click any result cell to see the formula in the formula bar.
5. Change the input values (e.g. marks, amounts, names) and watch the results update instantly.

> 💡 **Note:** `XLOOKUP` and `XMATCH` need Excel 2021 or Microsoft 365. Older versions can use `INDEX` + `MATCH` instead.

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 📁 Project Files

```
📦 Excel-Project
 ┣ 📄 pr1.xlsx      → Main workbook (3 sheets)
 ┣ 📄 README.md     → Project documentation
 ┗ 📁 assets
    ┣ header.svg    → Animated wave header
    ┣ typing.svg    → Typing text animation
    ┣ divider.svg   → Moving line divider
    ┗ footer.svg    → Animated footer
```

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 🔮 Future Improvements

- 📊 Add charts and a summary dashboard for each sheet
- 🎨 Add conditional formatting (grade colours, discount highlights)
- 🔽 Add dropdown lists (data validation) for region, product and department
- 🧮 Build Pivot Tables for region-wise and department-wise reports

<img src="assets/divider.svg" width="100%" alt="divider"/>

## 👨‍💻 Author

**Your Name**
📧 your.email@example.com · 🔗 [LinkedIn](https://linkedin.com/in/your-profile)

<img src="assets/divider.svg" width="100%" alt="divider"/>

<div align="center">

<img src="assets/footer.svg" alt="Thanks for visiting" width="100%"/>

</div>
