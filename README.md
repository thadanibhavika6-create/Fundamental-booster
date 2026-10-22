# Fundamental-booster

<div align="center">

<img src="https://capsule-render.vercel.app/api?type=waving&color=0:0b3d2a,50:107c41,100:33c481&height=260&section=header&text=FUNDAMENTAL%20BOOSTER&fontSize=50&fontColor=ffffff&stroke=0b3d2a&strokeWidth=1&animation=twinkling&fontAlignY=36&desc=Excel%20Formulas%20%7C%20Lookups%20%7C%20Analytics%20%7C%20Practice%20Project&descAlignY=58&descSize=19" width="100%"/>

<img src="https://readme-typing-svg.demolab.com/?lines=Master+Excel+formulas+step+by+step;Logical+%7C+Lookup+%7C+Statistical+%7C+Text+%7C+Date+functions;3+Sheets+%E2%80%A2+60+Records+%E2%80%A2+20%2B+Functions;Students+%E2%86%92+Sales+%E2%86%92+Employees+datasets;Beginner-friendly+Excel+practice+project&font=Fira+Code&center=true&width=700&height=45&color=107C41&vCenter=true&size=22&pause=1200"/>

<br/>

![Excel](https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Sheets](https://img.shields.io/badge/Worksheets-3-1f6feb?style=for-the-badge)
![Functions](https://img.shields.io/badge/Functions-20%2B-8957e5?style=for-the-badge)
![Level](https://img.shields.io/badge/Level-Beginner--Intermediate-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge)

</div>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Project%20Overview&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Project Overview"/>
</div>


**Fundamental Booster** is a hands-on Excel project (`pr1.xlsx`) that applies **logical, lookup, statistical, text and date functions** to three practical datasets — **Students, Sales and Employees**. Each sheet answers real business-style questions such as *"Who are the top students?"*, *"What are the total sales for a region?"* and *"What is an employee's salary?"* using clean, reusable formulas.

It's a great practice project for building strong **Excel fundamentals**: references, nested `IF`, lookups, conditional aggregates and text/date handling.

<img width="2880" height="3864" alt="Dashboard@2x (1)" src="https://github.com/user-attachments/assets/f5e50629-228f-4c7f-b3c6-e51bae0afc51" />

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Workbook%20Structure&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Workbook Structure"/>
</div>


| # | Sheet | Records | What it covers |
|:-:|---|:-:|---|
| 1 | 🎓 `students grade` | 20 | Marks, total, percentage, grade, top-student flag |
| 2 | 🛒 `sale data` | 20 | Products, regions, discounts, eligibility, sales lookups |
| 3 | 👥 `Employee data` | 20 | Departments, salaries, tenure, text & dynamic lookups |

All data lives in **Excel Tables** (`Table1`, `Table4`, `Table5`), so formulas stay readable and expand automatically.

### 1️⃣ `students grade`
| Column | Description |
|---|---|
| Student ID, Name | Unique ID and name of each student |
| Math, Science, English | Subject marks |
| Enrollment Date, Date of Birth | Date fields |
| Grade | A–F based on the average of the three subjects |
| Score | `Top Student` if Math and Science are both above 80, else `Normal` |
| Total marks, Percentage | `Math + Science + English` and `(Total / 300) × 100` |

### 2️⃣ `sale data`
| Column | Description |
|---|---|
| Sales ID, Product Code, Product | Sale and product identifiers |
| Region, Salesperson | Where and by whom the sale was made |
| Amount, Date | Sale value and sale date |
| Discount Rate | 20% (≥ 40,000), 10% (≥ 20,000), 5% (≥ 10,000), else No Discount |
| Discount Type | `Eligible` when the amount is 7,500 or more |
| Month | Month name extracted with `TEXT()` |

### 3️⃣ `Employee data`
| Column | Description |
|---|---|
| Employee ID, Name | Unique ID and name |
| Department | Finance, HR, IT or Marketing |
| Salary | Employee salary |
| Joining Date | Date of joining |
| Working Days | Days since joining using `DATEDIF` and `TODAY()` |

**Relationships between sheets:** each sheet is an independent dataset, linked by one theme — *turning raw rows into answers with formulas.*

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Features%20%2F%20Operations%20Covered&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Features / Operations Covered"/>
</div>


<table>
<tr>
<td width="50%" valign="top">

**🔹 Logical Functions**
- Grade (A–F) with nested `IF`
- Top-student flag with `IF` + `AND`
- Discount slabs with nested `IF`
- Eligibility check with `IF` + `OR`

</td>
<td width="50%" valign="top">

**🔹 Lookup Functions**
- `VLOOKUP` for names and prices by ID/code
- `INDEX` + `MATCH` two-way lookup
- `XLOOKUP` on unsorted data
- `XMATCH` to find a position in a list

</td>
</tr>
<tr>
<td width="50%" valign="top">

**🔹 Statistical Functions**
- `COUNTIF` — students above 50 in Math
- `AVERAGEIFS` — average of scores above 60
- `SUMIFS` — total sales by product and region

</td>
<td width="50%" valign="top">

**🔹 Text, Date & Advanced**
- `LEFT`, `FIND`, `UPPER`, `LOWER`, `TEXT`
- `TODAY`, `DATEDIF`, `MONTH`
- `INDIRECT` for dynamic ranges
- Relative vs absolute references (`C2` vs `$C$2`)

</td>
</tr>
</table>

### 🧮 Analysis Tasks Solved

| Sheet | Task | Function | Result |
|---|---|---|:-:|
| Students | Students scoring above 50 in Math | `COUNTIF` | **19** |
| Students | Average Math score above 60 | `AVERAGEIFS` | **82.26** |
| Students | Fetch a student's name from ID 1006 | `VLOOKUP` | **Student 6** |
| Sales | Total sales of Mouse in the East region | `SUMIFS` | **37,335** |
| Sales | Amount for product code PC110 | `VLOOKUP` | **35,185** |
| Sales | Person 4's sale in February | `INDEX` + `MATCH` | **5,264** |
| Sales | Salesperson with amount 35,611 | `XLOOKUP` | **Person 12** |
| Sales | Position of "Mouse" in the product list | `XMATCH` | **3** |

**🏅 Grading scale**

| Average | 90+ | 80–89 | 70–79 | 60–69 | Below 60 |
|---|:-:|:-:|:-:|:-:|:-:|
| Grade | 🥇 A | 🥈 B | 🥉 C | D | F |

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Tech%20Stack&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Tech Stack"/>
</div>


![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Tables](https://img.shields.io/badge/Excel_Tables-107c41?style=flat-square)
![Formulas](https://img.shields.io/badge/Formulas-33c481?style=flat-square)

> ⚠️ **Note:** `XLOOKUP` and `XMATCH` need **Excel 2021 or Microsoft 365**.
> In older versions, use `INDEX` + `MATCH` instead.

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=How%20to%20Run&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="How to Run"/>
</div>


1. Download or clone this repository.
2. Open `pr1.xlsx` in Microsoft Excel (2021 / Microsoft 365 recommended).
3. Switch between the three sheets using the tabs at the bottom.
4. Click any result cell to see its formula in the formula bar.
5. Change the input values (marks, amounts, names) and watch the results update instantly.

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Sample%20Insights%20You%20Can%20Derive&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Sample Insights You Can Derive"/>
</div>


- 🏅 Top-performing students and their grades
- 📈 Grade distribution across 20 students
- 💰 Total sales by product, region and salesperson
- 🏷️ Which sales qualify for a discount, and how much
- 👥 Salary lookup for any employee, in any order of data

<br/>

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=File%20Structure&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="File Structure"/>
</div>


```
📦 Fundamental-booster
 ┣ 📄 pr1.xlsx      # Main workbook (3 sheets)
 ┗ 📄 README.md     # Project documentation
```

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Future%20Improvements&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Future Improvements"/>
</div>


- [ ] Add charts and a summary dashboard for each sheet
- [ ] Add conditional formatting (grade colours, discount highlights)
- [ ] Add dropdown lists (data validation) for region, product and department
- [ ] Build Pivot Tables for region-wise and department-wise reports

<br/>

<div align="center">

<div align="center">
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=110&section=header&text=Author&fontSize=28&fontColor=ffffff&fontAlignY=38&animation=twinkling" width="100%" alt="Author"/>
</div>

**Bhavika Thadani**
📍 Ahmedabad

Made with ❤️ to strengthen Excel fundamentals — formulas, lookups & data analysis.

⭐ **If you found this project useful, consider giving it a star!**

<img src="https://capsule-render.vercel.app/api?type=waving&color=0:107c41,100:33c481&height=120&section=footer&text=Thanks%20for%20visiting!&fontSize=24&fontColor=ffffff&animation=twinkling&fontAlignY=65" width="100%"/>

</div>
