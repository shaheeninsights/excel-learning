# My Excel learning Log
This folder tracks my daily Excele Learning using Luke Barousse’s YouTube course.

## Day 1 - Setup + Excel basics

- Created this GitHub repo
- Connected my local Excel folder to Git
- Created the MyLearning folder for my own notes
- Used a fictitious dataset provided by the instructor
- Applied formulas with basic maths operators (e.g., +, -, *, /, ^, &).
- Learned the order of mathematical operator execution.
- Used comparison operators (=, >, <, >=, <=, <>).
- Applied cell referencing and autofill.
- Applied percentage format on a column.
- Added needed decimal places as per the requirement.
- Learned fixed and absolute referencing using the $ sign.

## Day 2 – Functions Intro

- Used the same fictitious dataset.
- Explored why functions are used in Excel.
- Calculated averages both manually by cell and using the AVERAGE() function with a cell range.
- Applied the AND() function to evaluate logical conditions.
- Explored the Formula tab, which provides explanations and descriptions of functions.
- Used COUNT() and COUNTIF() functions to count values based on conditions, practicing with static and dynamic criteria.
- Practiced combining static values and cell references using the & operator.
- Used the Mac shortcut Command + T to toggle absolute and relative cell references.

## Errors

Explored common Excel error types and how to handle them: #DIV/0! (e.g., =1/0) division by zero error; #VALUE! (e.g., =B4 + "text") wrong argument type in formula; #REF! (e.g., =#REF!) invalid cell reference, often due to deleted cells; #NAME? (e.g., =COUNTT(A3:A9)) unrecognized function or name; #N/A (e.g., =VLOOKUP("Value", A1:A10, 2, FALSE)) data not available for lookup; #NUM! (e.g., =SQRT(-1)) invalid numeric value in formula; #NULL! (e.g., =SUM(A1:A10 B1:B10)) Incorrect range intersection.

## Day 3 – Logical Functions, Math Functions & Funnel Chart

### Logical Functions
Using the same fictitious dataset, I practiced several logical functions to classify and filter data:

- **IF()** – returns values based on a condition  
- **AND()** – checks whether multiple conditions are TRUE  
- **OR()** – checks if at least one condition is TRUE  
- **IFS()** – evaluates multiple conditions for data bucketing  

These were used to analyse job titles, countries, and salary ranges.

### Math Functions
On a larger dataset, I practiced core math functions:

- **SUM()**, **AVERAGE()**, **MIN()**, **MAX()**, **COUNT()**  
- **COUNTIF()** and **AVERAGEIF()** for conditional calculations  

I created summary tables showing totals, averages, minimums, maximums, and counts for different job categories and locations.

### Funnel Chart
I created a funnel chart using the summary values, learning how to:

- Insert a chart  
- Select appropriate data  
- Format and present the chart clearly  

This helped reinforce how to visualise decreasing values across categories.

## Day 4 – Statistical Functions, Salary Analysis & Ranking

### Statistical Functions
Using a subset of the job postings dataset (three columns: job title, country, and average yearly salary), 
I practiced core statistical functions to summarise the data. These included:

- **COUNT()** – number of salary entries  
- **AVERAGE()** – mean salary  
- **MEDIAN()** – middle value in the salary distribution  
- **STDEV()** – standard deviation to measure salary spread  
- **MIN() / MAX()** – lowest and highest salaries  
- **QUARTILE()** – 1st, 2nd (median), and 3rd quartiles  
- **MODE()** – most frequently occurring salary value  

These functions helped me understand how salaries vary across roles and countries.

### Salary Analysis by Job Title
Next, I calculated the **average salary for each job title** in the dataset.  
This allowed me to compare roles such as:

- Data Scientist  
- Data Engineer  
- Data Analyst  
- Machine Learning Engineer  
- Business Analyst  
- Senior-level roles  

This step reinforced how grouping and aggregating data help reveal trends.

### Ranking Salaries
I then used the **RANK()** function to rank job titles from highest‑paid to lowest‑paid based on their average salaries.  
This created a clear, ordered list showing which roles earn the most.

### Visualisation
After sorting the ranked salaries, I inserted a **horizontal bar chart** (recommended chart).  
This chart visually shows:

- highest‑paid roles at the top  
- lowest‑paid roles at the bottom  
- all ranks clearly displayed  

This helped me understand how to turn summary tables into visual insights.

## Day 5 – Array Functions, Unique Values, Sorting, Median & Monthly Job Counts

### Array Functions
Today I practiced **Array Functions** using the same job‑posting dataset.  
I learned about two types of arrays:

- **Modern Dynamic Arrays** – spill automatically and update instantly  
- **Classic Arrays** – older Excel arrays that require Ctrl + Shift + Enter  

Dynamic arrays made it easier to extract, sort, and analyse multiple values at once.

### Dataset Columns Used
I worked with three columns from the dataset:

- `job_title_short`
- `job_posted_date`
- `salary_year_avg`

### Unique Job Titles
I extracted all unique job titles using the **UNIQUE()** function:  =UNIQUE(A2:A32673)

### Sorting Unique Job Titles
I sorted the unique job titles alphabetically using **SORT()**:  =SORT(R2#)

### Median Salary by Job Title
I calculated the **median salary** for each job title using a conditional array 
formula:  =MEDIAN(IF(($A$2:$A$32673=$S2)*($M$2:$M$32673<>""),$M$2:$M$32673))

This formula:
- checks if the job title matches  
- ignores blank salary values  
- returns the median salary for that job title  

### Number of Jobs Posted Each Month
I used **TEXT()** to extract the month name from the posted date, then counted how many jobs were posted each month using this 
formula: =SUMPRODUCT(--(TEXT($H$2:$H$32673,"mmmm")=$V2))

This allowed me to analyse job‑posting activity by month.

## Day 6 – Lookup Functions (VLOOKUP & XLOOKUP)

### Dataset Used
Today I worked with the Job Openings dataset.  
The columns used were:

- `job_title_short`
- `job_country`
- `salary_year_avg`

### VLOOKUP Practice
I used **VLOOKUP** to find the **company name** associated with:

- the **minimum** salary  
- the **maximum** salary  
- the **median** salary  

This helped me understand how VLOOKUP searches vertically through a table to return related information.

### XLOOKUP Practice
I then practiced **XLOOKUP**, which is more flexible and powerful than VLOOKUP.

I used XLOOKUP to return:

- the **job title** for the minimum, maximum, and median salary  
- the **country** for the minimum, maximum, and median salary  

This allowed me to pull information from any column without needing to count column numbers.

### Salary Bucketing with XLOOKUP
I also used **XLOOKUP** to categorise `salary_year_avg` into **salary buckets**.  
This involved mapping each salary to a bucket range (e.g., <75K, 75K–100K, 100K–125K, etc.) using a lookup table.

This exercise helped me understand how lookup functions can be used not only to retrieve data, but also to classify and organise it.

## Day 7 – Text Functions & Skill Extraction

### Dataset Used
Today I worked with the **Job Applicants dataset**, which contains columns such as:

- Applicant ID  
- Full Name  
- Job Position  
- Application Date & Time  
- Email  
- Street  
- City State Zip  
- Skills List  

This dataset was used to practice different **Text Functions** in Excel and to analyse applicant skills.

### 1. Text Combination (TEXTJOIN)
I used the **TEXTJOIN** function to combine address information into one complete address.

Columns used:
- Street  
- City State Zip  

These were joined into a single address field using `TEXTJOIN`, which helped me understand how to merge text from separate columns.

### 2. Text Extraction
I practiced extracting specific parts of text using different Excel functions:

- **TEXTSPLIT()**  
  - Used to split the *Full Name* column  
  - Example: extracting the **first name** from a full name

- **RIGHT()**  
  - Used to extract the **last 3 characters** from the Applicant ID  
  - Useful for identifying codes or patterns inside IDs

- **TEXTSEARCH(), FIND(), MID()**  
  - Used together to extract the **State** from the combined `City State Zip` field  
  - This showed how to locate text positions and pull out specific segments

These exercises helped me understand how to break down and manipulate text fields for cleaning and analysis.

### 3. Skill Extraction & Counting
At the end, I worked on identifying the **skills of all job applicants**.

Steps completed:
1. Used **TEXTJOIN** to combine all skills from the Skills List column into one long text string  
2. Used **TEXTSPLIT** to separate the combined skills into individual skills  
3. Used **TRANSPOSE** to convert the row of skills into a column  
4. Used **UNIQUE** to get a list of distinct skills  
5. Used **COUNTIF** to count how many applicants have each skill  

This produced a clean table showing each skill and how many applicants possess it.

### What the Skills Chart Shows
The final skills chart displays:

- The **most common skills** among all job applicants  
- Which skills appear **frequently** vs. **rarely**  
- A clear comparison of skill popularity across the applicant pool  

This helps identify:
- The top skills applicants commonly have  
- Skills that are less common and may represent gaps  
- Overall skill distribution in the dataset  

## Day 8 – Date and Time Functions

### Dataset Used
Today I worked with the **Job Posting dataset**, but only used the first **20 rows** because date and time functions can be time‑consuming on large datasets.

### 1. Date Functions
I practiced extracting different parts of a date using:

- **YEAR()** – extracts the year  
- **MONTH()** – extracts the month number  
- **DAY()** – extracts the day number  

I also used:

- **DATE()** – to construct a date from year, month, and day  
- **TODAY()** – to return the current date  
- **DATEDIF()** – to calculate the number of days since an application was made  

This helped me understand how to break down and work with date values for analysis.

### 2. Time Functions
I practiced extracting time components from the application timestamp using:

- **HOUR()** – extracts the hour  
- **MINUTE()** – extracts the minute  
- **SECOND()** – extracts the second  

I also noted that time can be extracted using the **TEXT()** function as an alternative method.

### 3. Hour‑Based Application Analysis
To analyse the time of day when applications are submitted, I used:

- **SEQUENCE()** to generate **24 rows**, representing each hour of the day  
- Extracted the hour from each application timestamp  
- Counted how many applications were made in each hour  

This allowed me to see the distribution of applications across the day.

### Final Chart
A chart was created using:

- **24 hours of the day**  
- **Number of applications submitted in each hour**  

### What the Chart Shows
The chart clearly shows:

- The **time of day when most applications are submitted**  
- In this dataset, the **highest number of applications were made at the end of the day**  
- Lower activity during early morning hours  
- A gradual increase as the day progresses  

This provides insight into applicant behaviour and peak activity times.

## Day 9 – Introduction to Charts

### Dataset Used
Today I worked with the **Job Posting dataset** to learn the basics of charts in Excel.  
Charts are also called **plots**, **graphs**, or **visualisations**.  
They are powerful because, as the saying goes, *a picture speaks a thousand words*.

Excel provides:
- **Recommended Charts** – helpful suggestions based on the data  
- **All Charts** – full list of chart types with customisation options  

I also explored chart elements such as:
- Axes  
- Axis titles  
- Chart title  
- Gridlines  
- Data labels  
- Legend  
- Trendline  
- Error bars  
- Up/Down bars  
- Quick Layout options  

---

## 1. Line Chart – Job Postings by Month
A **line chart** is best for **time‑series data** because it shows how values change over time and how each point is connected.

For this dataset, I created a line chart showing:
- **X‑axis:** Months  
- **Y‑axis:** Count of job postings  
- **Chart title:** Job Postings by Month  
- **Y‑axis title:** Count of Jobs  
- **Trendline:** Added to show the overall direction  

This chart helps identify seasonal patterns or monthly trends in job postings.

---

## 2. Pie Chart – Jobs Mentioning a Degree
A **pie chart** is used to show **proportions** or **percentages** of a whole.

I created a pie chart to answer the question:

**“What jobs mention a degree?”**

The chart showed:
- **Degree mentioned:** 81%  
- **No degree mentioned:** 19%  

This visualisation makes it easy to compare the share of postings that require a degree versus those that do not.

---

## 3. Column Chart – Job Count by Job Title
A **column chart** uses **vertical bars**.  
It is ideal for comparing values across categories.

Using the job posting dataset, the column chart showed:
- Data Analyst roles have the highest number of postings  
- Followed by Data Scientist and Data Engineer  
- Senior roles have significantly fewer postings  

This chart helps compare job demand across different job titles.

---

## 4. Bar Chart – Top Jobs in Data Science
A **bar chart** uses **horizontal bars**.  
It is useful when:
- Category names are long  
- You want an easy left‑to‑right comparison  
- You want to rank categories clearly  

The bar chart displayed:
- Data Analyst: 9.6K  
- Data Scientist: 8.5K  
- Data Engineer: 6.8K  
- Senior roles: much lower counts  

This chart clearly shows the ranking of job roles in terms of demand.

---

### Difference Between Column and Bar Charts
| Chart Type | Orientation | Best For |
|-----------|-------------|----------|
| **Column Chart** | Vertical bars | Comparing values when categories are short or few |
| **Bar Chart** | Horizontal bars | Long category names, ranking, easier readability |

Both charts compare categories, but the orientation changes readability and emphasis.

---

### Summary
Today’s work helped me understand:
- When to use each chart type  
- How to customise chart elements  
- How charts reveal insights quickly  
- How different chart types communicate different messages  

Charts make data easier to understand and are essential for presenting analysis clearly.

## Chart Summary Table

| Chart Type      | Orientation / Shape | Best Used For | Strengths | Example From Dataset |
|-----------------|---------------------|----------------|-----------|-----------------------|
| **Line Chart**  | Connected points forming a line | Time‑series data | Shows trends over time, highlights increases/decreases | Job postings by month with a trendline |
| **Pie Chart**   | Circular chart divided into slices | Showing proportions or percentages | Easy to compare parts of a whole | % of job postings mentioning a degree vs not |
| **Column Chart** | Vertical bars | Comparing categories when labels are short | Good for side‑by‑side comparisons, easy to read | Job count by job title (Data Analyst, Data Scientist, etc.) |
| **Bar Chart**   | Horizontal bars | Comparing categories with long labels or ranking | Best for readability when category names are long; great for ranking | Top jobs in data science (Data Analyst, Data Scientist, etc.) |

## Day 10 – Advanced Charts

### Dataset Used
Today I continued working with the **Job Posting dataset**, the same dataset used in Day 9.  
The focus was on deeper visual analysis of job pay and job location using **advanced chart types** and **customisation techniques**.

---

## 1. Scatter Plot – Comparing Yearly and Hourly Pay

A **scatter plot** is ideal when comparing **two numerical values**.  
In this analysis, I compared:

- `salary_year_avg`  
- `salary_hour_avg`  

This allowed me to see how yearly and hourly pay relate across different job titles.

### Customisation Practiced
I explored several advanced formatting options:

- Adjusting **axis bounds** to improve readability  
- Changing **number formats** (custom formatting for currency)  
- Applying a **clean font** to reduce visual clutter  
- Adding **axis titles**  
- Adding an appropriate **chart title**  
- Adding **data labels** to show the job title for each point  
- Using **leader lines** to keep labels neat and readable  
- Adding a **trendline** to show the overall relationship between hourly and yearly pay  

### What the Scatter Plot Shows
The scatter plot reveals:

- Senior roles (e.g., Senior Data Engineer, Senior Data Scientist) have the **highest yearly and hourly pay**  
- Entry‑level roles (e.g., Data Analyst, Business Analyst) cluster at the lower end  
- There is a clear **positive relationship** between hourly and yearly pay  
- Job titles are spread across the chart, showing variation in compensation across the data field  

This chart helps identify which roles offer the highest compensation and how pay scales across different job types.

---

## 2. Map Chart – Job Counts and Salary by Country

A **map chart** is useful for showing **geographical patterns** in data.

I created two map charts:

### a) Job Count by Country
This map shows where most jobs in the dataset are located.

Key insights:
- The United States has the **highest number of job postings**  
- Other countries such as Canada, the UK, India, and Australia also show notable job counts  
- Lighter shades indicate fewer postings, darker shades indicate more  

This helps identify global job distribution.

### b) Median Salary by Country
This map shows which countries offer the **highest median salaries**.

Key insights:
- Darker regions represent **higher median pay**  
- Some countries show significantly higher salary levels than others  
- This helps compare earning potential across different regions  

---

## Summary
Today’s work focused on:

- Using **advanced chart types** (scatter plot and map chart)  
- Customising charts for clarity and presentation quality  
- Analysing job pay across roles and countries  
- Understanding how geography and job title influence compensation  

These visualisations provide deeper insights into the dataset and help communicate findings more effectively.

## Day 11 – Charts & Statistics

### Dataset Used
Today I worked with the **salary dataset**, specifically focusing on **yearly salary data**.  
The goal was to understand statistical charts that help describe the **distribution**, **spread**, and **shape** of salary data.

Two major chart types were explored:
1. Histogram  
2. Box & Whisker Plot  

These charts are essential for understanding how salary values are distributed and where the data is concentrated.

---

## 1. Histogram – Understanding Salary Distribution

A **histogram** shows how many values fall into specific ranges.  
Each bar represents a **bin**, and each bin covers a salary range.

### What I Learned
- Each bar shows the **count of salaries** within that range  
- Bins are **equally spaced**  
- When one bin ends, the next begins — no overlap  
- Smaller bin width = more detail, but can look noisy  
- Larger bin width = smoother but less detailed  
- Axis limits were fixed to avoid sensory overload and make the chart readable  
- Titles and labels were added for clarity  

### What the Histogram Shows
The histogram revealed:

- Most Data Analyst salaries in the US fall between **$77K and $94K**  
- The next most common ranges are **$60K–$77K** and **$94K–$112K**  
- Very few jobs pay above **$129K**  
- The distribution is **right‑skewed**  
  - Meaning: a long tail on the right due to a few very high salaries  
  - This is common in salary data because a small number of roles pay extremely high amounts  

### Why Histograms Matter
Histograms help answer:
- Where do most salaries fall?  
- Are salaries evenly spread or concentrated?  
- Are there outliers pulling the distribution?  
- Is the data skewed left or right?  

This is the foundation of understanding **salary ranges and market expectations**.

---

## 2. Box & Whisker Plot – Understanding Spread, Median & Outliers

A **box and whisker plot** (boxplot) is a powerful statistical chart that summarises the entire dataset in one visual.

### How to Read a Boxplot
- The **box** shows the **Interquartile Range (IQR)**  
  - Left side of the box = **Q1 (25th percentile)**  
  - Right side of the box = **Q3 (75th percentile)**  
- The **line inside the box** = **Median (50th percentile)**  
- The **X inside the box** = **Mean (average)**  
- The **whiskers** show the range of values within 1.5 × IQR  
- Any dots beyond the whiskers = **Outliers**  
  - These are unusually high or low values  

### What the Boxplot Shows (Salary Data)
From the boxplots:

- The median salary sits well below the highest values  
- There are many **high‑salary outliers**  
- The IQR shows where the **middle 50%** of salaries lie  
- The whiskers show the typical salary range  
- Outliers represent very high‑paying roles  

### Why Boxplots Matter
Boxplots help answer:
- What is the median salary?  
- How spread out are salaries?  
- Are there extreme outliers?  
- Which job titles have the highest typical pay?  
- How do different job roles compare?  

### Boxplots by Job Title
When boxplots were paired with job titles (e.g., Data Scientist, Data Engineer, Senior roles):

- Senior roles had **higher medians** and **more outliers**  
- Entry‑level roles had **lower medians** and **tighter ranges**  
- Some roles (like Senior Data Engineer) had extremely high outliers  
- This makes it easy to compare pay across job categories  

---

## Summary
Today’s work focused on understanding **statistical charts** that describe how salary data behaves.

### Key Takeaways
- **Histograms** show the *shape* of the data — where values cluster and how they spread  
- **Boxplots** show the *summary statistics* — median, quartiles, range, and outliers  
- Salary data is often **right‑skewed** because a few roles pay extremely high salaries  
- Boxplots make it easy to compare salary distributions across job titles  

These charts help analysts understand not just the numbers, but the **story behind the numbers**.

## Day 12 – Sparklines

### Dataset Used
Today I continued working with the **Job Posting dataset**, focusing specifically on **monthly job posting counts** for different job roles.  
The goal was to learn how to use **sparklines**, which are tiny charts that sit inside a cell and summarise trends in a compact way.

---

## What Are Sparklines?
A **sparkline** is a small chart placed inside a single cell.  
It gives a quick visual summary of the data next to it — without taking up space like a full chart.

Sparklines are useful for:
- spotting trends  
- comparing patterns across rows  
- showing increases and decreases  
- highlighting seasonality  

They are perfect when you want a **quick visual snapshot** rather than a full chart.

---

## Types of Sparklines in Excel
Excel provides three sparkline types (Insert → Sparklines):

1. **Line Sparkline**  
   - Shows trends over time  
   - Best for monthly or sequential data  

2. **Column Sparkline**  
   - Shows highs and lows clearly  
   - Good for comparing values  

3. **Win/Loss Sparkline**  
   - Used for positive vs negative values  
   - Not needed for this dataset  

For this task, I used **Line Sparklines** because job postings are tracked **month by month**, and a line makes the trend easy to see.

---

## How Sparklines Were Created
Steps practised today:

1. Selected the **monthly job posting values** for each job role  
2. Insert → **Sparklines → Line**  
3. Choose the **location range** (the cell where the sparkline should appear)  
4. Applied custom formatting:
   - Changed **line colour**  
   - Highlighted **high points**  
   - Highlighted **low points**  
   - Adjusted **marker colours** to make the sparkline easier to read  

This made each sparkline visually clear and consistent.

---

## What the Sparklines Show
Each sparkline represents the **trend of job postings across 12 months** for one job role.

From the sparklines:

- **Data Analyst** and **Data Scientist** roles show strong peaks mid‑year  
- Senior roles have **lower counts** but follow similar seasonal patterns  
- Some roles show dips around October–November  
- The **Total** row sparkline shows the overall trend across all roles combined  

Sparklines make it easy to compare:
- Which roles are growing  
- Which roles are stable  
- Which roles have seasonal fluctuations  

All without needing a full chart.

---

## Why Sparklines Are Useful
Sparklines are powerful because they:

- Save space  
- Provide instant visual insight  
- Work well in dashboards  
- Help compare multiple rows quickly  
- Make tables more meaningful  

They are especially useful when dealing with **time‑series data** like monthly job postings.

---

## Summary
Today’s work focused on using **sparklines** to visualise monthly job posting trends.  
I learned how to insert, customise, and interpret sparklines, and how they help summarise data patterns directly inside a table.

Sparklines are a great tool for quick insights and are commonly used in dashboards and reports.

## Day 13 – Spreadsheet Advance

### Dataset Used
Today I worked with a **subset of the salary dataset**, focusing on advanced spreadsheet features that improve structure, automation, and analysis.  
This is the final chapter of Excel Basics and introduces tools that make data handling more efficient and professional.

---

## 1. Introduction to Tables

### Creating a Table
To convert a normal range into a table:
- Click anywhere inside the dataset  
- Go to **Insert → Table**  
- Excel automatically selects the full range  
- Ensure **“My table has headers”** is checked  

This creates a structured table with built‑in formatting and functionality.

### Exploring the Table Design Tab
The **Table Design** tab includes several useful options:

- **Table Name**  
  - Rename the table to something simple and meaningful  
  - Example: `SalaryTable`  
  - Makes formulas easier to read and reference  

- **Table Style Options**  
  - Header Row  
  - First Column  
  - Last Column  
  - Banded Rows  
  - Filter Buttons  
  These help improve readability and navigation.

### Benefits of Using Tables
Tables provide several advantages:

- **Automatic expansion**  
  - Adding a new column automatically fills formulas down the entire column  

- **Structured References**  
  - Instead of cell references like A2:A1000, Excel uses names like:  
    `=SalaryTable[@YearlySalary]`  
  - The `@` symbol refers to the current row  
  - This makes formulas easier to understand and maintain  

- **Easy copying**  
  - Typing `=SalaryTable` in another sheet copies the entire table  

- **Special reference options**  
  - `#Headers` → header row  
  - `#Data` → only the data  
  - `#All` → entire table  
  - You can also reference specific columns  

These features make tables ideal for dashboards, automation, and clean reporting.

### Limitations of Tables
- Large tables can slow down older computers  
- Some advanced Excel features don’t work directly with tables  
- But overall, tables are extremely useful for structured analysis

---

## 2. Subtotal and Aggregate Functions

### SUBTOTAL
The **SUBTOTAL** function can:
- Sum  
- Count  
- Average  
- Max  
- Min  
- And more  

Its main advantage:
- It **ignores filtered‑out rows**, making it perfect for summarising data inside tables.

### AGGREGATE
The **AGGREGATE** function is more advanced:
- Can ignore errors  
- Can ignore hidden rows  
- Supports 19 different operations  

Useful when working with messy or incomplete datasets.

---

## 3. Slicers for Tables

### What Are Slicers?
A **slicer** is a visual filter that allows you to filter a table by clicking buttons instead of using dropdown menus.

### How to Insert a Slicer
- Select the table  
- Go to **Table Design → Insert Slicer**  
- Choose the column to filter by  

### Customising Slicers
- Change the **caption** to make it more readable  
- Adjust the **style** for better visibility  
- Enable **multi‑select** to filter multiple categories at once  

Slicers make filtering:
- Faster  
- Cleaner  
- More interactive  
- Ideal for dashboards and presentations  

### Excel Table Formula Formats (Structured References)

When a normal range is converted into a **Table**, Excel uses *structured references* instead of cell references.  
These make formulas easier to read, understand, and maintain.

Below are the most important formats used inside Excel Tables:

---

## 1. Referencing the Entire Table
`=TableName`

Example:  
`=SalaryTable`  
Returns the full table including headers and data.

---

## 2. Referencing Specific Columns
`=TableName[ColumnName]`

Example:  
`=SalaryTable[YearlySalary]`  
Returns the entire YearlySalary column.

---

## 3. Referencing Headers Only
`=TableName[#Headers]`

Example:  
`=SalaryTable[#Headers]`  
Returns only the header row.

---

## 4. Referencing Data Only (No Headers)
`=TableName[#Data]`

Example:  
`=SalaryTable[#Data]`  
Returns only the data rows.

---

## 5. Referencing the Entire Table (Headers + Data)
`=TableName[#All]`

Example:  
`=SalaryTable[#All]`  
Useful when copying the entire table to another sheet.

---

## 6. Referencing a Column Within the Current Row
`=TableName[@ColumnName]`

Example:  
`=SalaryTable[@YearlySalary]`  
Returns the YearlySalary value for the current row only.

This is one of the most powerful features of tables — formulas automatically fill down.

---

## 7. Referencing Multiple Columns
`=TableName[[Column1]:[Column2]]`

Example:  
`=SalaryTable[[JobTitle]:[YearlySalary]]`  
Returns a range from JobTitle to YearlySalary.

---

## 8. Combining Row and Column References
`=TableName[@[Column1]:[Column2]]`

Example:  
`=SalaryTable[@[MinSalary]:[MaxSalary]]`  
Returns the range of values for that row only.

---

## 9. Using Structured References Inside Functions
Structured references work inside all Excel functions.

Examples:

**SUM of a column:**  
`=SUM(SalaryTable[YearlySalary])`

**AVERAGE of a column:**  
`=AVERAGE(SalaryTable[YearlySalary])`

**IF using row reference:**  
`=IF(SalaryTable[@YearlySalary] > 100000, "High", "Low")`

---

## Why Structured References Matter
- No more confusing cell references  
- Formulas auto‑fill and auto‑adjust  
- Easier to read and debug  
- Perfect for dashboards and reports  
- Makes your Excel work look professional  

Structured references are one of the biggest advantages of using tables in Excel.

## Day 14 – Formatting & Conditional Formatting

### Dataset Used
Today I continued working with the **Job Posting dataset**, focusing on formatting techniques and conditional formatting.  
The goal was to make data easier to read, interpret, and visually analyse.

---

## 1. Basic Cell Formatting

Before applying conditional formatting, I cleaned up the sheet:

- Used **Editing → Clear → Clear Formats** to remove old formatting  
- Converted the range into a **Table** for cleaner structure  
- Applied **Cell Styles → Heading 2** for the table title  
- Used **Merge & Center** to centre the table title  
- Adjusted column widths and alignment for readability  

These steps ensure the sheet looks clean and professional before adding visual rules.

---

## 2. Introduction to Conditional Formatting

**Conditional Formatting** allows Excel to highlight or format cells *dynamically* based on rules.  
This makes patterns easier to spot without manually checking values.

Examples:
- Highlighting high salaries  
- Colour‑coding job counts  
- Showing icons for rankings  
- Adding data bars to show magnitude  

Conditional formatting is found under:
**Home → Conditional Formatting**

---

## 3. Colour Scales

Colour scales apply a gradient based on the value in each cell.

- Green → higher values  
- Yellow → mid‑range  
- Red → lower values  

This is useful for quickly spotting:
- Which job titles have the highest salaries  
- Which roles have the lowest job counts  
- Where WFH% is high or low  

I also practiced:
- **Clear Rules → Selected Cells**  
- **Clear Rules → Entire Sheet**  
to remove formatting when needed.

---

## 4. Format Painter

The **Format Painter** helps copy formatting from one cell or range to another.

Steps:
1. Select a formatted cell  
2. Click **Format Painter**  
3. Click the target cell(s)  

This keeps formatting consistent across the sheet.

---

## 5. Managing Rules

Using **Conditional Formatting → Manage Rules**, I explored:

- Viewing all rules applied to the worksheet  
- Editing rule order  
- Changing rule types  
- Applying rules to specific ranges  

This is essential when multiple rules overlap or when formatting becomes complex.

---

## 6. Data Bars (Applied to Job Count Column)

For the **Job Count** column, I used **Data Bars**.

Why data bars?
- They visually show the magnitude of each value  
- Perfect for comparing counts  
- Less distracting than colour scales  
- Easy to read horizontally  

Customisation:
- Chose a subtle colour to avoid visual overload  
- Ensured bars were readable but not overpowering  
- Applied only to the first column to keep the table clean  

---

## 7. Icon Sets (Ratings)

Next, I created a **New Formatting Rule** using **Icon Sets**.

Steps:
1. Conditional Formatting → New Rule  
2. Select **Icon Set**  
3. Choose a rating style (e.g., stars, flags, arrows)  
4. Tick **“Show Icon Only”** to hide the numbers  
5. Adjust thresholds (e.g., top 20%, mid 50%, bottom 30%)  

This created a clean, visual ranking system for job roles.

### Why Icon Sets Are Useful
- They summarise performance at a glance  
- They reduce clutter  
- They work well for dashboards  
- They make comparisons intuitive  

I also compared:
- A **clean, subtle icon‑only version**  
- A **distracting, overly colourful version**  

This helped me understand the importance of choosing formatting that supports the data instead of overwhelming it.

---

## 8. Good vs Bad Conditional Formatting

Today I learned that conditional formatting can either:

### ✔ Enhance the data  
- Subtle colours  
- Clean icon sets  
- Simple data bars  
- Consistent formatting  

### ✘ Or distract from the data  
- Too many colours  
- Multiple overlapping rules  
- Bright, clashing formats  
- Icons on every column  

The goal is to **support the analysis**, not overpower it.

---

## Summary
Today’s work focused on improving the visual clarity of the dataset using:

- Basic formatting  
- Conditional formatting  
- Colour scales  
- Data bars  
- Icon sets  
- Rule management  
- Format painter  

These tools help make data more readable, highlight important patterns, and create professional‑looking spreadsheets.

## Day 15 – Collaboration & Data Validation

### Dataset Used
Today I continued working with the **salary dataset**, focusing on two important Excel skills used in real workplaces:
1. Collaboration (protecting sheets and dashboards)
2. Data Validation (controlling user input)

These tools help prevent mistakes, protect dashboards, and ensure clean, reliable data entry.

---

## 1. Collaboration: Protecting Sheets & Workbooks

When sharing files with coworkers or stakeholders, it’s important to protect:
- formulas  
- dashboards  
- lookup sheets  
- data sources  

This prevents accidental edits that could break the entire workbook.

### Protecting a Sheet
Steps:
1. Go to **Review → Protect Sheet**  
2. Add a password (optional but recommended)  
3. Choose what users are allowed to do (e.g., select cells, sort, filter)

### Why Protect Sheets?
- Prevents accidental deletion of formulas  
- Keeps dashboards intact  
- Ensures coworkers only edit the intended cells  
- Protects lookup tables and data sources  

### Hiding Supporting Sheets
Often dashboards rely on:
- lookup tables  
- helper sheets  
- data sources  

To avoid confusion:
- Right‑click sheet → **Hide**  
- Protect workbook structure if needed  

This ensures users only see the dashboard and input sheet, not the backend logic.

---

## 2. Data Validation – Controlling User Input

Data validation restricts what users can type into a cell.  
This prevents:
- spelling mistakes  
- invalid job titles  
- wrong salary inputs  
- broken formulas  

### Example Used Today: Basic Salary Calculator

I created a simple calculator where the user selects a **job title** from a dropdown, and Excel returns the **median salary**.

This prevents users from typing:
- wrong job titles  
- misspellings  
- values that don’t exist in the dataset  

### Steps to Build the Calculator

#### **Step 1: Create a List of Unique Job Titles**
On a sheet named **data_validation**:
- Extracted unique job titles from the main dataset  
- Also added job count for reference  

This list becomes the source for the dropdown.

#### **Step 2: Apply Data Validation**
On the **basic calculator** sheet:
1. Select the cell where the job title will be chosen  
2. Go to **Data → Data Validation**  
3. Choose **List**  
4. Set the **Source** to the unique job titles range  

Now the user can only pick from the dropdown — no typing allowed.

---

## 3. Returning the Median Salary

To calculate the median salary based on the selected job title, I created a new sheet called **median_salary**.

This sheet contains:
- Job titles  
- Their median salaries  
- Sorted and cleaned for lookup  

### Formula Used
Two approaches were practiced:

#### **1. MEDIAN + IF (Array Formula)**
Used to calculate the median salary for each job title:
=MEDIAN(IF(job_title_range = selected_title, salary_range))
(Entered as a dynamic array formula)

#### **2. XLOOKUP**
Used in the calculator to return the median salary:
=XLOOKUP(selected_job_title, job_title_list, median_salary_list)


This makes the calculator dynamic and user‑friendly.

---

## 4. Protecting the Calculator

Once the calculator was working:
- Locked all formula cells  
- Left only the dropdown cell unlocked  
- Applied **Review → Protect Sheet**

This ensures:
- Users can only select a job title  
- They cannot break formulas  
- The calculator remains functional and clean  

---

## Workbook Structure Used Today

The workbook now contains four sheets:

1. **Data**  
   - Original salary dataset  

2. **Basic Calculator**  
   - Dropdown job title  
   - Median salary result  

3. **Data Validation**  
   - Unique job titles  
   - Job counts  

4. **Median Salary**  
   - Job titles  
   - Calculated median salaries  
   - Used for XLOOKUP  

This structure keeps everything organised and easy to maintain.

---

## Summary
Today’s work focused on collaboration and data validation — essential skills for real‑world Excel use.

### Key Takeaways
- Protect sheets to prevent accidental edits  
- Hide backend sheets to keep dashboards clean  
- Use data validation to control user input  
- Build dropdown‑based calculators for clean interaction  
- Use XLOOKUP and MEDIAN to return dynamic results  
- Lock formulas and protect sheets before sharing  

These techniques ensure your Excel files are professional, reliable, and safe to share with others.

## Day 16 – Salary Dashboard (Setup & Repo Initialization)

### Dataset Used
Today I began working on the **Salary Dashboard project**, using the dataset provided by the course instructor, **Luke Barousse**.  
This marks the start of the dashboard section, where all previous Excel skills will be applied in a real‑world project.

---

## 1. Reviewing the Project README
I started by going through the project’s **README file** to understand:

- The structure of the dashboard  
- The metrics that will be displayed  
- The sheets required  
- The workflow expected in the project  
- The final deliverables  

This gave me a high‑level overview of what the dashboard will look like and how the data will be used.

---

## 2. Initialising the Git Repository

To keep the project organised and version‑controlled, I:

1. Created a **local Git repository** on my machine  
2. Added the project files  
3. Made the initial commit  
4. Pushed the repository to **GitHub**  

This ensures:
- My work is backed up  
- I can track changes  
- I can share the project easily  
- I follow good data‑analysis workflow practices  

---

## 3. Project Structure (Initial)
At this stage, the project contains:

- The dataset  
- The README  
- The initial Excel workbook  
- The Git repo connected to GitHub  

More sheets, calculations, and visuals will be added as the dashboard develops.

---

## Summary
Today was the setup day for the Salary Dashboard project.  
I reviewed the project instructions, prepared my environment, and pushed the initial version to GitHub.  
The next steps will involve cleaning the data, building supporting sheets, and starting the dashboard layout.

## Day 17 – Salary Dashboard (Country Dropdown Setup)

### Dataset Used
Continuing with the job posting dataset from Luke Barousse’s Excel course, today I worked on building the interactive components of the Salary Dashboard. The focus was on creating the **Country** dropdown filter, which will later be used to calculate median salaries and update dashboard visuals dynamically.

---

## Setting Up the Dashboard Sheet
I began by preparing the layout of the dashboard:

- Positioned the main dashboard title correctly.
- Ensured spacing and alignment matched the intended structure.
- Confirmed the existing **Job Title** dropdown was functioning.
- Identified where the new **Country** dropdown should be placed.

This ensures the dashboard remains clean and consistent as more filters and charts are added.

---

## Creating the Country Data Source
To support the Country dropdown, I created a new backend sheet named **country**, which will later store:

- `job_country`
- `median_salary` (used for map chart and lookups)

This sheet will act as a structured reference for future formulas and visuals.

---

## Preparing the Data Validation Source
All dropdowns must come from a controlled list.  
On the **data_validation** sheet, I created a new column for unique country names.

### Extracting Unique Countries
To pull all distinct countries from the dataset:=UNIQUE(jobs[job_country])

This ensures the dropdown only includes valid country names.

### Sorting the List
To make the dropdown easier to use: =SORT(G2#)

This sorted list was labelled **job_country_sorted**.

Sorting improves usability and keeps the dashboard consistent with the instructor’s version.

---

## Adding the Country Dropdown to the Dashboard
With the sorted list ready, I added the dropdown to the dashboard.

### Steps:
1. Select the cell where the Country dropdown will appear.
2. Go to **Data → Data Validation**.
3. Choose **List**.
4. Set the **Source** to the sorted country list on the data_validation sheet.
5. Apply the rule.

The dashboard now displays a clean, dynamic dropdown of all countries in the dataset.

---

## How This Fits Into the Full Dashboard
According to the project README, the final dashboard will include:

- Job Title filter  
- Country filter  
- Job Type filter  
- A multi‑criteria MEDIAN formula  
- A bar chart for job salaries  
- A map chart for country median salaries  

Today’s work completes the **second major filter** (Country), which is required for:

- The map chart  
- The multi‑criteria median salary calculation  
- Dynamic updates to dashboard visuals  

---

## Summary
Today I completed the **Country dropdown** for the Salary Dashboard by:

- Structuring the dashboard layout  
- Creating a backend sheet for country data  
- Extracting unique country names  
- Sorting the list for usability  
- Linking the sorted list to the dashboard via data validation  

This prepares the dashboard for the next steps, where additional filters and formulas will be added to make the dashboard fully interactive.

## Day 18 – Salary Dashboard (Job Type Dropdown Setup)

### Dataset Used
Continuing with the job posting dataset from Luke Barousse’s Excel course, today’s focus was on creating the **Job Type** dropdown for the Salary Dashboard. This is the third major filter (after Job Title and Country) and will later be used in the multi‑criteria MEDIAN formula that powers the dashboard.

---

## Setting Up the Type Sheet
To keep the project organised, I created a new sheet named **Type**.  
This sheet will eventually store:

- `job_schedule_type`
- `median_salary` (used later for filtering and calculations)

For now, the main goal was to extract and clean the unique job schedule types from the dataset.

---

## Preparing the Data Validation Source (Job Type)
The raw dataset contains job schedule types in the column `jobs[job_schedule_type]`.  
However, this column contains messy values such as:

- Multiple schedule types combined with **“and”**
- Comma‑separated values
- Occasional blanks or zeros

These must be cleaned before creating a dropdown.

### Step 1 — Extract Unique Job Types
On the **data_validation** sheet, I added a new column: =UNIQUE(jobs[job_schedule_type])

This returns all distinct schedule types, but the list still contains combined values like:

- “Full‑time and Contract”
- “Part‑time, Internship”

These cannot be used directly in a dropdown.

---

## Cleaning the Job Type List
To clean the list, I needed to remove:

- Any value containing **“and”**
- Any value containing **commas**
- Any zero or blank values

### Step 2 — Identify values containing “and”
Using the SEARCH function:=SEARCH("and", J2#)


This returns the position of the word “and” if it exists.

### Step 3 — Convert to TRUE/FALSE
Wrapping SEARCH inside ISNUMBER:=ISNUMBER(SEARCH("and", J2#))


- TRUE → “and” exists  
- FALSE → clean value  

### Step 4 — Filter out unwanted values
Using FILTER to keep only clean job types:

=FILTER(
J2#,
(NOT(ISNUMBER(SEARCH("and", J2#)))) *
(NOT(ISNUMBER(SEARCH(",", J2#)))) *
(J2# <> 0)
)

This formula removes:

- Values containing “and”
- Values containing commas
- Zero or blank entries

The result is a clean list of valid job schedule types.

### Step 5 — Sort the Clean List
To make the dropdown user‑friendly: =SORT(K2#)

This sorted list was named **job_schedule_type_sorted**.

---

## Adding the Job Type Dropdown to the Dashboard
With the cleaned and sorted list ready, I added the dropdown to the dashboard.

### Steps:
1. Select the cell where the Job Type dropdown will appear.
2. Go to **Data → Data Validation**.
3. Choose **List**.
4. Set the **Source** to the sorted job schedule type list on the data_validation sheet.
5. Apply the rule.

The dashboard now displays a clean dropdown of valid job schedule types.

---

## Why This Step Matters
The Job Type dropdown is essential for:

- The multi‑criteria MEDIAN formula  
- Filtering the bar chart  
- Filtering the map chart  
- Ensuring the dashboard responds to all three user selections:
  - Job Title  
  - Country  
  - Job Type  

Cleaning the job schedule types ensures the dashboard does not break due to inconsistent or combined values.

---

## Summary
Today I completed the **Job Type dropdown** for the Salary Dashboard by:

- Creating a new Type sheet  
- Extracting unique job schedule types  
- Cleaning the list using SEARCH, ISNUMBER, and FILTER  
- Sorting the cleaned list  
- Linking the sorted list to the dashboard via data validation  

This completes all three dropdown filters required for the interactive dashboard.















