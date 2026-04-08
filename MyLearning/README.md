# My Excel learning Log
This folder tracks my daily Excel Learning using Luke Barousse’s YouTube course.

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

## Adding the Job Type Dropdown to the Dashboard
With the cleaned and sorted list ready, I added the dropdown to the dashboard.

### Steps:
1. Select the cell where the Job Type dropdown will appear.
2. Go to **Data → Data Validation**.
3. Choose **List**.
4. Set the **Source** to the sorted job schedule type list on the data_validation sheet.
5. Apply the rule.

The dashboard now displays a clean dropdown of valid job schedule types.

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

## Summary
Today I completed the **Job Type dropdown** for the Salary Dashboard by:

- Creating a new Type sheet  
- Extracting unique job schedule types  
- Cleaning the list using SEARCH, ISNUMBER, and FILTER  
- Sorting the cleaned list  
- Linking the sorted list to the dashboard via data validation  

This completes all three dropdown filters required for the interactive dashboard.

## Day 19 – Salary Dashboard (Country Chart Setup)

### Dataset Used
Continuing with the job posting dataset from Luke Barousse’s Excel course, today I focused on building the **Country Chart** for the Salary Dashboard. This chart will visually display **median salaries by country**, based on the filters selected by the user.

## 1. Preparing the Country Sheet

To support the chart, I worked on the **country** sheet, which acts as the backend data source.

### Step 1 — Pulling Country Names
I used the following formula to bring in the sorted list of countries from the data_validation sheet: =Data_Validation!H2#

This ensures the country list is consistent with the dropdown filter used on the dashboard.

---

## 2. Calculating Median Salary by Country

For each country, I calculated the **median salary** using a multi‑criteria array formula:

=MEDIAN(
IF(
(jobs[job_country]=A2) *
(jobs[salary_year_avg]<>0) *
(jobs[job_title_short]=title) *
(jobs[job_schedule_type]=type),
jobs[salary_year_avg]
)
)


### Explanation:
- Filters the dataset by:
  - Selected country (`A2`)
  - Non-zero salaries
  - Selected job title
  - Selected job type
- Returns the **median salary** for matching rows

This formula powers the map chart and ensures the salary data reflects all three filters.

## 3. Handling #NUM! Errors

Some countries returned `#NUM!` errors in the median salary column.  
This happens when:

- No matching rows exist for the selected filters  
- The filtered salary list is empty  
- MEDIAN cannot calculate a result from an empty array  

To clean this up, I created a new column called **job_country_filter** using: =SORT(FILTER(A2:B112, ISNUMBER(B2:B112)), 2, -1)


### What this does:
- Filters out rows where the median salary is not a number  
- Sorts the remaining rows by salary (descending)  
- Returns a clean list of countries and their valid median salaries

This filtered list is used as the source for the map chart.

## 4. Building the Country Chart

With the filtered data ready, I inserted a **Map Chart** on the dashboard:

- Selected the **job_country_filter** column and corresponding **median_salary**  
- Inserted the chart using **Insert → Maps → Filled Map**  
- Positioned the chart below the Country dropdown

### Why a Map Chart?
- Visually shows salary differences across countries  
- Color-coded regions make comparisons intuitive  
- Matches the instructor’s final dashboard layout

## Summary

Today I built the **Country Chart** for the Salary Dashboard by:

- Pulling a clean list of countries from the validation sheet  
- Calculating median salary using a multi‑criteria formula  
- Filtering out invalid results using ISNUMBER + FILTER  
- Sorting the valid results for chart readability  
- Inserting a map chart to display global salary trends

This chart responds dynamically to the selected filters and helps users understand geographic salary differences.

## Day 20 – Salary Dashboard (Highlighting Selected Job in Bar Chart)

### Dataset Used
Continuing with the job posting dataset from Luke Barousse’s Excel course, today I worked on building the **Job Title Salary Bar Chart** and adding dynamic highlighting so the selected job title stands out visually on the dashboard. This makes the dashboard more intuitive and helps users quickly compare their chosen job against others.

## 1. Preparing the Title Sheet for Chart Data
To support the bar chart, I expanded the **title** sheet by adding helper columns. These columns calculate:

- Median salary for each job title  
- Sorted job titles  
- Highlighting logic for the selected job  

This sheet acts as the backend data source for the bar chart.

## 2. Calculating Median Salary for Each Job Title
In column **B**, I used a multi‑criteria MEDIAN formula:
                                                        =MEDIAN(
IF(
(jobs[job_title_short]=A2) *
(jobs[salary_year_avg]<>0) *
(jobs[job_country]=country) *
(ISNUMBER(SEARCH(type, jobs[job_schedule_type]))),
jobs[salary_year_avg]
)
)

### What this formula does:
- Filters the dataset by:
  - Job title (A2)
  - Non‑zero salaries
  - Selected country
  - Selected job type (using SEARCH to match partial text)
- Returns the **median salary** for that job title under the selected filters

This ensures the bar chart updates dynamically when the user changes dropdown selections.

## 3. Sorting Job Titles by Median Salary
To make the bar chart easier to read, I sorted the job titles based on their median salary:=SORT(A2#: B2#, 2, 1)

### Explanation:
- Sorts the job title list (A2#) and salary list (B2#)
- Sorts by the **second column** (median salary)
- Sorts in **descending order** (1)

This sorted list is used directly in the bar chart.

## 4. Creating Highlighting Logic for the Chart
To highlight the selected job title in the bar chart, I created two new columns:

### Column F — All other jobs (light blue)
=IF($D2<>title, $E2, NA())

### Column G — Selected job only (dark blue)
=IF($D2=title, $E2, NA())

### Why this works:
- Column F contains salary values for **all jobs except the selected one**
- Column G contains salary **only for the selected job**
- NA() hides the bar in the chart
- When both series are plotted:
  - Column F = **light blue** bars  
  - Column G = **dark blue** bar (highlighted)

This creates a clean visual effect where the selected job title stands out clearly.

## 5. Building the Bar Chart on the Dashboard
With the backend data ready, I inserted a **horizontal bar chart** on the dashboard:

- Series 1 → Column F (light blue, non‑selected jobs)
- Series 2 → Column G (dark blue, selected job)
- Category labels → Sorted job titles

### Formatting applied:
- Removed chart border
- Adjusted bar thickness
- Applied custom number formatting to the salary axis:$#,##0,"K"

This displays salaries like:
- $90K  
- $125K  
- $155K  

This matches the instructor’s dashboard style.

## 6. Result on the Dashboard
The bar chart now:

- Updates automatically when the user changes:
  - Job Title  
  - Country  
  - Job Type  
- Highlights the selected job title in **dark blue**
- Shows all other job titles in **light blue**
- Displays salaries in a clean, readable format

This makes the dashboard more interactive and visually intuitive.

## Summary
Today I completed the **Job Title Salary Bar Chart** with dynamic highlighting by:

- Expanding the title sheet with helper columns  
- Calculating median salary using a multi‑criteria formula  
- Sorting job titles by salary  
- Creating highlight logic using IF + NA()  
- Building a two‑series bar chart with light and dark blue colors  
- Applying custom number formatting for readability  

This feature helps users instantly see how their selected job compares to others in terms of salary.

## Day 21 – Salary Dashboard (Job Type Chart & Backend Logic)

### Dataset Used
Continuing with the job posting dataset from Luke Barousse’s Excel course, today I built the **Job Type Salary Chart**. This chart shows how median salaries vary across different job schedule types (Full‑time, Contractor, Part‑time, Internship, Temp work) based on the user’s selected filters. To support this, I expanded the **type** sheet with new formulas, sorting logic, and highlighted columns.

## 1. Preparing the Type Sheet Structure

The type sheet now contains the following columns:

- **Column A** — job_schedule_type  
- **Column B** — median salary for each type  
- **Column C** — empty spacer column (intentional)  
- **Column D** — sorted job types  
- **Column E** — sorted median salaries  
- **Column F** — non‑selected type values (light blue bars)  
- **Column G** — selected type values (dark blue bar)  

This layout keeps raw data, sorted data, and highlight logic clearly separated and easy to maintain.

## 2. Pulling Job Schedule Types from Data Validation

I pulled the cleaned list of job schedule types from the Data_Validation sheet: =Data_Validation!K2

This ensures the type list matches the dropdown used on the dashboard.

## 3. Calculating Median Salary for Each Job Type

In column **B**, I used a multi‑criteria MEDIAN formula with slight adjustments from previous sheets:
=MEDIAN(
IF(
(jobs[job_title_short]=title) *
(jobs[salary_year_avg]<>0) *
(jobs[job_country]=country) *
(ISNUMBER(SEARCH(A2, jobs[job_schedule_type]))),
jobs[salary_year_avg]
)
)

### What this formula does:
- Filters the dataset by:
  - Selected job title  
  - Non‑zero salaries  
  - Selected country  
  - Selected job schedule type (A2)  
- Uses `SEARCH` to match partial text  
- Returns the **median salary** for that job type  

This ensures the chart updates dynamically when the user changes filters.

## 4. Sorting Job Types by Median Salary

To prepare the data for charting, I sorted the job types by their median salary: =SORT(FILTER(A2:B6, ISNUMBER(B2:B6)), 2, 1)

### Why this formula is needed:
- **FILTER(A2:B6, ISNUMBER(B2:B6))** removes rows where the median salary is not a number (e.g., no matching data for that country).
- **SORT(..., 2, 1)** sorts the filtered list by the salary column in **descending order**.

This ensures the chart displays job types from highest to lowest salary and avoids errors caused by #NUM! values.


## 5. Creating Highlighting Logic for the Chart

To highlight the selected job type in the chart, I added two helper columns:

### Column F — Non‑selected types (light blue)  =IF($D2<>type, $E2, NA())

### Column G — Selected type only (dark blue)  =IF($D2=type, $E2, NA())


### Why this works:
- Column F shows salaries for all types **except** the selected one  
- Column G shows salary **only** for the selected type  
- `NA()` hides the bar in the chart  
- When plotted together:
  - Column F = light blue bars  
  - Column G = dark blue bar  

This creates a clean visual highlight effect.

## 6. Building the Job Type Chart

I inserted a **clustered bar chart** using:

- **Series 1** → Column F (light blue)  
- **Series 2** → Column G (dark blue)  
- **Category labels** → Column D (sorted job types)  

### Formatting applied:
- Removed chart border  
- Adjusted bar thickness  
- Applied custom number formatting to the X‑axis: $#,##0,"K"


This displays salaries like:
- 97K  
- 122K  
- 130K  

This keeps the dashboard consistent with the Job Title chart.

## Summary

Today I completed the **Job Type Salary Chart** by:

- Pulling job schedule types from the validation sheet  
- Calculating median salary using a multi‑criteria formula  
- Filtering and sorting the results  
- Creating highlight logic using IF + NA()  
- Building a two‑series bar chart with light and dark blue colors  
- Applying custom number formatting for readability  

This chart now updates dynamically based on the user’s selected job title, country, and job type, making the dashboard more interactive and insightful.

## Day 22 – KPI Cards (Median Salary, Top Platform, Job Count)

### Overview
Today I added three KPI cards to the dashboard:
1. **Median Salary**
2. **Top Job Platform**
3. **Job Count**

Each KPI required backend logic, new formulas, and small adjustments to the data validation sheet. These KPIs help the dashboard communicate key insights at a glance.

## 1. KPI: Median Salary

### a. Pulling the median salary from the Title sheet
The median salary for the selected job title and country had already been calculated in the **title** sheet (column E after sorting). To bring this value into a KPI card:

In the **title** sheet, cell **I2**: =XLOOKUP(title, D2:D11, E2:E11)

- `title` is the selected job title from the dashboard.
- `D2:D11` contains sorted job titles.
- `E2:E11` contains the corresponding median salaries.

I renamed **I2** as: median_salary

This makes it easy to reference directly from the dashboard.

### b. Displaying the KPI on the dashboard
On the dashboard:
- Inserted a **text box**.
- In the formula bar, linked it to: =median_salary
- Applied formatting:
  - Currency
  - No decimal places
  - Bold, large font for visibility

This creates a clean, dynamic KPI card that updates whenever the user changes filters.=

## 2. KPI: Top Job Platform

This KPI identifies the platform where the selected job title appears most frequently (e.g., Indeed, LinkedIn, ZipRecruiter).

### a. Creating the Platform sheet
Created a new sheet named **platform**.

In **A1**:job_via

In **A2**:=UNIQUE(jobs[job_via])

This generates a list of all platforms where jobs were posted.

### b. Counting job occurrences per platform
To calculate how many jobs appear on each platform under the selected filters, I adapted the median salary logic into a **COUNT** formula.

In **B2**:
=COUNT(
IF(
(jobs[job_country]=country) *
(jobs[job_title_short]=title) *
(ISNUMBER(SEARCH(type, jobs[job_schedule_type]))) *
(jobs[job_via]=A2),
jobs[salary_year_avg]
)
)

### What this formula does:
- Filters the dataset by:
  - Selected country  
  - Selected job title  
  - Selected job type  
  - Platform (A2)
- Counts how many matching rows exist.

This gives the number of job postings per platform.

### c. Sorting platforms by job count
To identify the top platform:

In **D2**: =SORT(A2:B594, 2, -1)

This produces a clean platform name for display on the dashboard.

### e. Adding the KPI to the dashboard
- Inserted a text box.
- Linked it to the cleaned top platform cell.
- Applied formatting to match the dashboard style.

---

## 3. KPI: Job Count

This KPI shows how many jobs match the selected filters.

### a. Updating Data_Validation sheet
To support this KPI, I added logic to count jobs based on:
- Selected job title  
- Selected country  
- Selected job type  

In **Data_Validation!M2**: =XLOOKUP(title, D2:D11, E2:E11)


Then renamed **M2** as:=count
- Applied formatting for consistency.

---

## Summary

Today, I completed the backend logic and dashboard display for three KPI cards:

- **Median Salary** — pulled via XLOOKUP and formatted for display  
- **Top Job Platform** — built using UNIQUE, COUNT, SORT, and SUBSTITUTE  
- **Job Count** — added via XLOOKUP and linked to the dashboard  

These KPIs make the dashboard more informative and allow users to quickly understand salary levels, job availability, and where roles are most frequently advertised.

## Day 23 – Dashboard Formatting, Final Touch‑Up & Sheet Protection

### Overview
Today was all about polishing the dashboard: removing visual noise, applying consistent formatting, and protecting the workbook so users can interact with filters without accidentally breaking formulas or backend logic. These final touches make the dashboard feel professional, intentional, and ready for real‑world use.

## 1. Removing Gridlines for a Clean Dashboard

Gridlines make sense when building formulas, but they distract from a finished dashboard.  
To clean up the view:

**View → Untick “Gridlines”**

This instantly gives the dashboard a clean, modern look.

## 2. Cleaning Up Chart Formatting

Each chart needed a final visual pass to match the dashboard style.

### Steps:
- Selected each chart
- **Format → Shape Outline → No Outline**
- Ensured consistent font sizes and colors
- Checked spacing and alignment

Removing outlines helps the charts blend naturally into the dashboard instead of looking boxed‑in.

## 3. Applying Proper Headings Using Cell Styles

To keep typography consistent:

**Home → Cell Styles → Heading**

This ensures all section titles follow the same formatting rules (font, size, weight), which improves readability and visual hierarchy.

## 4. Hiding Backend Sheets

To prevent users from navigating into backend logic:

- Right‑click each backend sheet (jobs, title, type, platform, data_validation)
- **Hide**

Only the dashboard remains visible, giving a clean user experience.

## 5. Protecting the Dashboard Sheet

The goal is to allow users to interact with dropdowns and slicers **without** being able to edit formulas or layout.

### Steps:

1. **Turn Headings On**  
   View → Tick “Headings”  
   (This makes it easier to select all cells.)

2. **Select All Cells**  
   Click the small triangle in the top‑left corner of the grid.

3. **Unlock Only the Cells Users Should Interact With**  
   - Right‑click → Format Cells → Protection  
   - Untick **Locked**  
   (This is done only for dropdown cells.)

4. **Protect the Sheet**  
   Review → Protect Sheet  
   Allow only:
   - **Select unlocked cells**

   Everything else stays locked.

5. **Turn Headings Off Again**  
   View → Untick “Headings”

This ensures the dashboard is safe from accidental edits while still being fully interactive.

## Summary

Today I completed the final polish and protection steps for the dashboard:

- Removed gridlines for a clean visual layout  
- Removed chart outlines and applied consistent styling  
- Used cell styles for headings  
- Hid all backend sheets  
- Protected the dashboard while keeping dropdowns usable  

These finishing touches make the dashboard feel professional, stable, and ready for presentation or portfolio use.

# 📘 Day 24 — Pivot Tables in Excel

## 1. Introduction to Pivot Tables

A **pivot table** is one of Excel’s most powerful tools for summarising large datasets. It allows you to quickly **sort, count, total, and analyse** data without writing any formulas — everything is done through simple drag‑and‑drop fields.

If you have thousands of rows and want answers like *“Total sales by region”* or *“Count of job titles”*, a pivot table can produce it in seconds.

### What does “Pivot” mean?

The word **pivot** means to *rotate around a central point*.  
A pivot table does exactly that — it lets you **rotate your data view**, reorganising the same dataset to see it from different angles.

Same data → different perspective.

### Simple Example

Raw data:

| Salesperson | Region | Month | Sales |
|------------|--------|--------|--------|
| Ahmed | North | Jan | 500 |
| Sara | South | Jan | 300 |
| Ahmed | North | Feb | 700 |
| Sara | South | Feb | 400 |

Pivot table summary:

| Region | Jan | Feb | Total |
|--------|------|------|--------|
| North | 500 | 700 | 1200 |
| South | 300 | 400 | 700 |

You’ve **pivoted** the data — same rows, new insight.

### Why Pivot Tables Matter

- **Fast** — summarise thousands of rows instantly  
- **Flexible** — drag fields to change the layout  
- **Insightful** — totals, averages, comparisons, trends  
- **Safe** — original data stays untouched  
- **Business‑ready** — used across finance, HR, sales, marketing  

### Core Idea (One Sentence)

> A pivot table lets  **rearrange and summarise data** to view it from any angle — without writing formulas.

---

## 2. Creating a Pivot Table (Job Count Analysis)

For this task, I used the **Salary Dataset** to calculate how many times each job title appears in the data.

### Step 1 — Insert a Pivot Table

From the **Insert** tab, Excel provides two options:

- **PivotTable**  
- **Recommended PivotTables**

**Recommended PivotTables** shows pre‑built layouts and lets you choose whether to place the pivot table in the **existing sheet** or a **new sheet**.

**PivotTable** (manual option) is used when you already know what you want to analyse — in this case, the **count of each job title**.

### Step 2 — Choose the Data Source

Excel gives two choices:

1. **Select a table or range**  
2. **Use an external data source**

I selected **“Table/Range”** and chose to place the pivot table on a **new worksheet**.  
After it was created, I renamed the sheet to **job_count**.

### Step 3 — Build the Pivot Table

In the **PivotTable Fields** pane:

- Drag **job_title_short** into **Rows**  
- Drag **job_title_short** again into **Values**  

Excel automatically summarises it as **Count of job_title_short**, giving the job count for each title.

I also explored adding **job_country** to the **Filters** area to filter job counts by country.

---

## Summary

Today I learned how pivot tables help summarise large datasets quickly and flexibly. I created a pivot table to calculate job counts, explored the field layout, and used filters to refine the analysis. This is a key Excel skill for real‑world data analysis.

# 📘 Day 25 — Pivot Table Analysis & Monthly Salary Insights

## Overview

Today I continued working with PivotTables to deepen my understanding of Excel’s analytical features. My main focus areas were:

- Exploring PivotTable Analyse and Design tools  
- Calculating **average salary** (hourly and yearly) for different job roles  
- Building a **Monthly Job Count** PivotTable using the Salary Dataset  
- Using text functions to extract month names dynamically  
- Refreshing PivotTables to incorporate new calculated fields  

---

## 1. Renaming and Exploring PivotTable Analyse Tools

I started by renaming my existing PivotTable sheet to **job_count** for clarity.

Then I explored the **PivotTable Analyse** tab, including:

- **Active Field**  
- **Group**  
- **Filter**  
- **Data**  
- **Calculations**

This helped me understand how to manage fields, refresh data, and control summarisation behaviour.

---

## 2. Adding Average Salary Calculations

Next, I added two new measures to the PivotTable:

- **Average Hourly Salary**  
- **Average Yearly Salary**

After inserting these fields, I applied appropriate number formatting to make the values easier to read. This step helped me compare salary levels across different job roles more effectively.

---

## 3. Exploring PivotTable Design Options

I moved on to the **Design** tab to improve the visual layout of the PivotTable. I explored:

- **Report Layout**  
- **Subtotals and Grand Totals**  
- **PivotTable Style Options**  
- **Built‑in PivotTable Styles**

This allowed me to create a cleaner, more readable table suitable for analysis and presentation.

---

## 4. Final Analysis: Monthly Job Count by Job Title

My main objective was to create a PivotTable showing **job titles (rows)** and their **monthly posting counts (columns)**.

### Step 1 — Convert Data to a Table  
I converted the raw dataset into an Excel Table and named it **jobs**.  
This ensures the PivotTable updates automatically when new data is added.

### Step 2 — Insert a New PivotTable  
I inserted a PivotTable based on the **jobs** table and placed it on a **new worksheet**.

### Step 3 — Rename the Sheet  
I renamed the sheet to **Monthly Count** for clarity.

### Step 4 — Extract Month Names  
To analyse job postings by month, I created a new calculated column using: =TEXT([@[job_posted_date]], "mmm")


This extracted the month abbreviation (Jan, Feb, Mar, etc.) from each posting date.

After refreshing the PivotTable, the new **month** field appeared automatically — a great example of how dynamic PivotTables are.

### Step 5 — Clean Up the Layout  
To improve readability, I:

- Turned **Field Headers** off  
- Removed **Row Labels** and **Column Labels**  
- Added **job_country** to the Filters area to analyse specific regions  

### Result  
This produced a fully customisable PivotTable showing monthly job posting counts for each job title.

I added the final output below for reference.

---

## Final Output: Monthly Job Count Table

*(Displayed in the screenshot I added to this entry.)*

---

## Summary

Today I strengthened my PivotTable skills by:

- Adding calculated fields for average salary  
- Exploring Analyse and Design tools  
- Creating a dynamic month-based PivotTable  
- Refreshing PivotTables to incorporate new fields  
- Building a clean, filterable Monthly Job Count report  

This was a productive session that helped me understand how PivotTables can transform raw data into meaningful insights.

# 📘 Day 26 — Advanced Pivot Table Features (Hierarchy & Grouping)

## Overview

Today I explored **advanced PivotTable features** using the same Salary Dataset (30k+ rows). My focus was on understanding how Excel handles **hierarchies**, **automatic grouping**, and **manual grouping** inside PivotTables. These features help me analyse large datasets more efficiently and reveal deeper patterns in the data.

---

## 1. Pivot Table Hierarchy

### Step 1 — Insert PivotTable
I inserted a new PivotTable based on the **jobs** table and placed it on a new sheet.  
I renamed this sheet **Hierarchy**.

### Step 2 — Building the Hierarchy
My goal was to analyse **average salary by job title within each country**.

To create this structure:

- I placed **job_country** in **Rows**  
- I placed **job_title_short** under it in **Rows**  
- This automatically created a **parent → child hierarchy**  
  - Parent: Country  
  - Child: Job Titles  

### Step 3 — Adding Salary and Job Count
Next, I added:

- **salary_year_avg** → Values (Average)  
- **job count** → Values (Count)

This allowed me to see both the **average salary** and the **number of job postings** for each country and job title.

### Step 4 — Improving Layout
I switched to the **Design** tab and selected:

- **Report Layout → Show in Outline Form**

This made the hierarchy clearer and easier to read.

### Step 5 — Cleaning the Data
Some job titles had blank salary values, so I used:

- **Value Filters → Greater Than → 0**

This removed empty or invalid salary rows.

---

## 2. Automatic Grouping (Date Hierarchy)

Next, I explored how Excel automatically groups dates inside PivotTables.

### Step 1 — Insert PivotTable for Monthly Job Count
I created another PivotTable on a new sheet and named it **Group_Automatic**.

### Step 2 — Using job_posted_date
I dragged **job_posted_date** into **Rows**.

Excel automatically created a **date hierarchy**:

- **Month**  
  - **Days**  
    - **Times of day**  

This showed how Excel breaks down dates into multiple levels without manual work.

### Step 3 — Counting Job Posts
I added **job_title_short** into **Values** (Count).  
This gave me:

- Job postings per month  
- Job postings per day  
- Job postings per specific time of day  

### Step 4 — Exploring Drill‑Down
By double‑clicking any number, Excel opened the underlying rows used to calculate that value.  
This helped me understand exactly how the PivotTable aggregated the data.

---

## Summary

Today I learned how to use advanced PivotTable features to analyse large datasets more effectively:

- I created **hierarchies** to compare salaries and job counts by country and job title.  
- I used **automatic date grouping** to break down job postings by month, day, and time.  
- I improved PivotTable layouts using the **Design** tab.  
- I applied **value filters** to remove empty or invalid data.  

These features make PivotTables far more powerful for real‑world data analysis and reporting.

# 📘 Day 27 — Manual Grouping in PivotTables (Job Title Categorisation)

## Overview

Today I focused on **manual grouping** inside PivotTables to organise job titles into meaningful categories. This technique helps me simplify large datasets by creating custom parent groups, making it easier to analyse job distributions, percentages, and rankings.

---

## 1. Creating Custom Job Title Groups

I grouped the job titles into three logical categories:

### **Data Nerds**
- Data Analyst  
- Data Scientist  
- Data Engineer  

### **Senior Data Nerds**
- Senior Data Engineer  
- Senior Data Scientist  
- Senior Data Analyst  

### **Other Data Nerds**
- Business Analyst  
- Machine Learning Engineer  
- Software Engineer  
- Cloud Engineer  

These groups allow me to analyse the dataset at a higher level instead of reviewing each job title individually.

---

## 2. Building the PivotTable

### Step 1 — Insert PivotTable  
I created a new PivotTable from the existing **jobs** table and placed it on a new sheet.

### Step 2 — Add Job Titles  
I dragged **job_title_short** into the **Rows** area.

### Step 3 — Manual Grouping  
To create each group:

1. I selected the job titles I wanted to group  
2. Right‑clicked → **Group**  
3. Excel created a default group name (e.g., *Group1*)  
4. I renamed it in the formula bar (e.g., *Data Nerds*)  

I repeated this process for all three groups.

### Step 4 — What Excel Creates in the Background  
Excel automatically generated a new field called **job_title_short2**, which contains the **parent group names**:

- Data Nerds  
- Senior Data Nerds  
- Other Data Nerds  

This field becomes essential for calculating percentages and ranking.

---

## 3. Adding Job Counts and Percentages

### **Job Count**
I added **job_title_short** to the **Values** area and summarised it by **Count**.

### **Percentage of Grand Total**
I added the same field again and used:

- **Show Values As → % of Grand Total**

This shows how much each job title contributes to the entire dataset.

### **Percentage of Parent**
Using **job_title_short2** as the base, I applied:

- **Show Values As → % of Parent**

This shows how much each job title contributes **within its group**.

For example:  
If Data Analyst is 38.60% of the Data Nerds group, that means it accounts for 38.60% of all roles inside that category.

---

## 4. Adding Ranking

I added another calculation using:

- **Show Values As → Rank Largest to Smallest**

### **What Rank Represents**
Rank assigns a number based on job count:

- **Rank 1** = highest job count  
- **Rank 2** = second highest  
- **Rank 3** = third highest  

### **How Excel Calculates Rank**
Excel sorts the job counts from largest to smallest and assigns ranking values accordingly.  
If two job titles have the same count, Excel assigns the same rank and skips the next number (standard competition ranking).

### **Why Ranking Is Useful**
Ranking helps me quickly identify:

- The most in‑demand roles  
- The least common roles  
- How roles compare within their group  

This adds an extra layer of insight to the grouped PivotTable.

---

## 5. Final Output

The final PivotTable displays:

- Custom job groups  
- Job counts  
- % of Grand Total  
- % of Parent  
- Rank  

This gives me a clear, structured view of how different job titles contribute to the overall dataset and to their respective categories.

---

## Summary

Today I learned how to:

- Use **manual grouping** to create custom job categories  
- Understand how Excel generates a new grouping field  
- Calculate **job counts**, **% of total**, **% of parent**, and **rank**  
- Build a clean, hierarchical PivotTable for deeper analysis  

Manual grouping is a powerful technique for simplifying large datasets and creating meaningful categories that support better analysis.

# 📘 Day 28 — PivotCharts, Slicers & Timelines (Answering Key Analytical Questions)

## Overview

Today I explored **PivotCharts**, **Slicers**, and **Timelines** to answer three important analytical questions using the Salary Dataset. These tools help me turn PivotTables into interactive visual insights.

The three questions I aimed to answer were:

1. **How are jobs trending over time?**  
2. **Which job has the highest percentage of demand?**  
3. **What is the top‑paying job in data science?**

Each PivotChart below is built specifically to answer one of these questions.

---

# 1. PivotChart: Top‑Paying Job  
### **Q3 — What is the top‑paying job in data science?**

To answer this question, I needed to calculate the **average yearly salary** for each job title and visualise the results.

### Step 1 — Create PivotTable  
I inserted a new PivotTable using the **jobs** table.

### Step 2 — Add Fields  
- Dragged **job_title_short** into **Rows**  
- Dragged **salary_year_avg** into **Values**  
- Changed summarisation from **Sum → Average**

### Step 3 — Format Salary  
I formatted the values as **Currency** with **0 decimal places**.

### Step 4 — Rename Column  
I renamed the values field to **Average Yearly Salary**.

### Step 5 — Insert PivotChart  
Insert → PivotChart → **Column Chart**

### Step 6 — Sort  
Sorted **Average Yearly Salary** from **Largest → Smallest** to instantly reveal the highest‑paying roles.

**This PivotChart directly answers Q3 by showing which job title has the highest average yearly salary.**

---

# 2. PivotChart + Slicers: Highest Demand Job  
### **Q2 — Which job has the highest percentage of demand?**

To answer this question, I used the PivotTable from **Group_Manual**, which already contains:

- Job counts  
- % of Grand Total  
- Grouped job titles (Data Nerds, Senior Data Nerds, Other Data Nerds)

### Step 1 — Clean the PivotTable  
I removed all other value fields except **% of Grand Total**, since this is the metric needed to measure demand.

### Step 2 — Insert PivotChart  
Insert → PivotChart → **Column Chart**

### Step 3 — Add Slicers  
To make the chart interactive:

PivotChart Analyse → **Insert Slicer**

I added slicers for:

- **job_title_short2** (the grouped job categories)  
- **job_country** (to filter demand by region)

### Step 4 — Format Slicers  
I customised:

- Slicer captions  
- Multi‑select options  
- Layout and style  

**This PivotChart answers Q2 by showing which job title contributes the highest percentage to the overall job market.**

---

# 3. PivotChart + Timeline: Job Counts by Month  
### **Q1 — How are jobs trending over time?**

To answer this question, I used the PivotTable created earlier for **Monthly Job Count**, which summarises job postings by month.

### Step 1 — Insert PivotChart  
Insert → PivotChart → **Line Chart**

### Step 2 — Add Trendline  
Design → Add Chart Element → **Trendline → Linear**

This helps visualise the overall direction of job posting trends.

### Step 3 — Insert Timeline  
PivotChart Analyse → **Insert Timeline**

Excel detected **job_posted_date**, so I selected it.

### Step 4 — Explore Timeline Features  
I tested:

- Switching between **Months**, **Quarters**, and **Years**  
- Selecting multiple periods  
- Formatting the timeline  
- Renaming the slicer caption to **Date**

### Step 5 — Connect Filters  
PivotTable Analyse → **Filter Connections**  
I selected the PivotTables that should respond to the timeline.

**This PivotChart answers Q1 by showing how job postings rise or fall over time, with the timeline allowing month‑by‑month or quarter‑by‑quarter exploration.**

---

# Summary

Today I learned how to:

- Build PivotCharts for salary, demand, and time‑based analysis  
- Use **Slicers** to create interactive filters for job groups and countries  
- Use **Timelines** to filter charts by month, quarter, or year  
- Add **trendlines** to show long‑term patterns  
- Format and sort PivotCharts for clearer insights  

These tools make dashboards more dynamic and user‑friendly, allowing deeper exploration of the Salary Dataset.

# 📘 Day 29 — Advanced Data Analytics: Analysis Add-Ins & Data Tables

## Overview

Today I explored **Analysis Add-Ins** and **Data Tables** as part of advanced data analytics in Excel. These tools extend Excel's built-in capabilities and allow for powerful what-if analysis and scenario modelling.

The key topics covered were:

1. **What are Excel Add-Ins?**
2. **Analysis ToolPak Add-In**
3. **Forecast Sheet (Windows only — workaround used on Mac)**
4. **What-If Analysis**
5. **One-Input Data Tables**
6. **Two-Input Data Tables**

---

## What are Excel Add-Ins?

**Add-Ins** are optional tools that extend Excel's functionality beyond its default capabilities — similar to installing apps on a smartphone.

There are two types:

- **Built-in Add-Ins** — already inside Excel but switched off by default (e.g. Analysis ToolPak, Solver, Power Pivot)
- **Third-party Add-Ins** — downloaded from the Microsoft Store or other sources (e.g. Copilot, Power BI)

### How to Enable Add-Ins on Mac

On **Windows**: File → Options → Add-Ins

On **Mac** (equivalent path):
- For built-in Add-Ins → **Tools** (top Mac menu bar) → **Excel Add-Ins** → tick the Add-In → OK
- For Store Add-Ins → **Insert** tab → **Get Add-Ins**

> ⚠️ **Mac Note:** The "File → Options" path shown in the tutorial does not exist on Mac. The equivalent on Mac is Tools → Excel Add-Ins from the top menu bar of the screen (not inside the Excel ribbon).

---

## 1. Analysis ToolPak

The **Analysis ToolPak** is a built-in Add-In that provides advanced statistical and analytical tools — without needing to write complex formulas manually.

Once enabled it adds a **Data Analysis** button to the **Data** tab.

It includes tools for:
- Descriptive statistics
- Histograms
- Regression analysis
- Correlation calculations
- Moving averages

### How to Enable on Mac
1. Click **Tools** in the top Mac menu bar  
2. Click **Excel Add-Ins**  
3. Tick **Analysis ToolPak**  
4. Click **OK**  
5. A **Data Analysis** button now appears in the **Data** tab  

---

## 2. Forecast Sheet

The **Forecast Sheet** is a visual tool that takes historical time-series data and automatically generates a forecast chart with confidence bounds.

> ⚠️ **Mac Note:** Forecast Sheet is **not available on Mac** — it is a Windows-only feature. As a workaround on Mac, the same result can be achieved using formulas:

### Mac Workaround — Using Formulas

The dataset used was `Forecast_Original` — containing two columns:
- **Date** — daily dates throughout the year  
- **Job Count** — number of jobs posted each day  

**Step 1 — Extend dates** below the last existing row for future periods  

**Step 2 — Forecast column:**: =FORECAST.ETS(A367, $B$2:$B$366, $A$2:$A$366)

**Step 3 — Lower Confidence Bound:**: =FORECAST.ETS(A367,$B$2:$B$366,$A$2:$A$366) - FORECAST.ETS.CONFINT(A367,$B$2:$B$366,$A$2:$A$366)

**Step 4 — Upper Confidence Bound:**: =FORECAST.ETS(A367,$B$2:$B$366,$A$2:$A$366) + FORECAST.ETS.CONFINT(A367,$B$2:$B$366,$A$2:$A$366)


**Step 5** — Select all four columns → Insert → **Line Chart**

This produces the same chart as the Windows Forecast Sheet showing actual data, forecast line, and upper/lower confidence bounds.

### Key Concepts Learned

**Seasonality** — a repeating pattern in data at regular intervals.  
In this job posting dataset the pattern was **weekly** — high job counts on weekdays (Monday–Friday) and very low counts on weekends (Saturday–Sunday). This repeats every 7 days throughout the entire year.

**Why FORECAST.ETS instead of FORECAST.LINEAR:**

| | FORECAST.LINEAR | FORECAST.ETS |
|---|---|---|
| Seasonality | Ignores it | Accounts for it |
| Best for | Simple steady trends | Data with repeating cycles |

**Moving Average** — smooths out short-term daily fluctuations to reveal the underlying long-term trend.  
A **7-day moving average** is ideal for this dataset because each calculation covers exactly one full week (5 weekdays + 2 weekend days), cancelling out the weekly spikes and drops to produce a clean trend line.

---

## 3. What-If Analysis

**What-If Analysis** allows you to change input values and see how those changes affect your results — without manually editing formulas one by one.

The dataset used (`What-If_Analysis` sheet) modelled salary growth over 5 years with three input cells:

| Input | Value |
|---|---|
| Base Salary | £100,000 |
| Bonus | 10% |
| Annual Raise | 1.5% |

The result cells calculated the salary amount for each of Years 0–4 and a Total.

Excel's What-If Analysis tools are found under:  
**Data tab → What-If Analysis**

Three tools are available:
- **Scenario Manager** — save and compare multiple sets of input values  
- **Goal Seek** — work backwards from a desired result to find the required input  
- **Data Tables** — test multiple values of one or two inputs simultaneously  

---

## 4. One-Input Data Table

A **One-Input Data Table** tests how changing **one input variable** affects one or more result cells — all in one structured table, without manually changing the input each time.

The dataset (`Data_Table_One-Input` sheet) tested how different **Annual Raise** percentages affect salary across all 5 years and the total.

Annual raise values tested: 0%, 0.5%, 1%, 1.5%, 2%, 2.5%, 3%, 3.5%, 4%

Each row showed the resulting salary for Years 0–4 and Total for that specific raise percentage — giving an instant side-by-side comparison.

### How to Create a One-Input Data Table

1. Set up your input variable values in a **column**  
2. Place your result formula references in a **row** at the top  
3. Select the entire table range (inputs + results area)  
4. Go to **Data tab → What-If Analysis → Data Table**  
5. Leave **Row input cell** blank  
6. Set **Column input cell** to your input variable cell (e.g. Annual Raise cell)  
7. Click **OK**  

Excel fills in all results automatically.

---

## 5. Two-Input Data Table

A **Two-Input Data Table** tests how changing **two input variables simultaneously** affects a single result cell.

The dataset (`Data_Table_Two-Input` sheet) extended the salary model to test combinations of different **Annual Raise** percentages and different **Bonus** percentages against the Total salary.

This creates a grid where:
- **Rows** represent one variable (e.g. different bonus values)  
- **Columns** represent another variable (e.g. different raise values)  
- Each cell in the grid shows the result for that specific combination  

### How to Create a Two-Input Data Table

1. Place one variable's values in a **column**  
2. Place the other variable's values in a **row**  
3. Place your **single result formula** at the intersection of the row and column headers  
4. Select the entire table range  
5. Go to **Data tab → What-If Analysis → Data Table**  
6. Set **Row input cell** to the variable in your row  
7. Set **Column input cell** to the variable in your column  
8. Click **OK**  

### Key Difference from One-Input

| | One-Input | Two-Input |
|---|---|---|
| Variables tested | 1 | 2 |
| Result cells | Multiple | Only 1 |
| Output | Multiple columns of results | Single grid of results |

---

## Summary

Today I learned how to:

- Understand and enable **Excel Add-Ins** on Mac via Tools → Excel Add-Ins  
- Use **FORECAST.ETS** and confidence bound formulas as a Mac alternative to Forecast Sheet  
- Understand **seasonality** and **moving averages** in time-series data  
- Use **What-If Analysis** to model salary scenarios  
- Build **One-Input Data Tables** to test a single variable across multiple values  
- Build **Two-Input Data Tables** to test two variables simultaneously against one result  

These tools are essential for **scenario modelling and data-driven decision making** — allowing rapid testing of assumptions without manually changing formulas one by one.

# 📘 Day 30 — Analysis ToolPak: Descriptive Statistics, Histograms, Ranking, Percentile & Moving Averages

## Overview

Today I explored the **Analysis ToolPak** using the job postings salary dataset (30k+ rows).  
This Add‑In provides advanced statistical tools that allow me to perform deeper analysis without writing formulas manually.

The four areas I focused on were:

1. **Descriptive Statistics** on the salary column  
2. **Customisable Histograms**  
3. **Ranking and Percentile** of salaries  
4. **Moving Average** calculations  

These tools help automate exploratory data analysis (EDA) and create more interpretable visualisations.

---

## 1. Descriptive Statistics (EDA on Salary Data)

The **Descriptive Statistics** tool provides instant summary metrics such as mean, median, standard deviation, range, skewness, kurtosis, and more — without writing any formulas.

### Steps I followed:

1. On the **Data** tab, I opened **Data Analysis**, which displays a dialog box containing various statistical tools.
2. I selected **Descriptive Statistics**, which opened another dialog box.
3. For the input range, I selected the **salary_year_avg** column.
4. I checked **Labels in first row**.
5. I ticked **Summary Statistics** and left all other options as default.
6. After clicking **OK**, Excel generated a new sheet containing all descriptive metrics.
7. I applied some formatting to make the output easier to read.

This gives a complete statistical overview of the salary distribution without needing to manually calculate each metric.

---

## 2. Histogram (Visualising & Customising Salary Data)

Creating a histogram directly using **Insert → Chart → Histogram** results in unreadable x‑axis labels because the salary range is too wide.

To fix this, I used the **Histogram tool** inside the Analysis ToolPak, which allows full control over bin sizes.

### Steps I followed:

1. Data → **Data Analysis** → **Histogram**
2. Input Range → **salary_year_avg**
3. Checked **Labels**
4. Left all other settings as default  
   → The output was still not readable because the bins were too broad.

### Creating Custom Bins

To improve readability:

1. I created a new sheet and manually built a **Bin** column.
2. I used a **£50k bin size**, dragging it down until **£400k**.
3. I copied the **Bin** and **Frequency** labels to prepare the structure.
4. I went back to **Data Analysis → Histogram**.
5. This time, I selected the **custom bin range** I created.

This produced a much clearer histogram where each bin represented a meaningful salary range.

---

## 3. Rank & Percentile (Ranking Salaries)

To analyse how salaries compare across the entire dataset, I used the **Rank & Percentile** tool inside the Analysis ToolPak.

### Steps I followed:

1. Data → **Data Analysis**
2. Selected **Rank and Percentile**
3. Input Range → **salary_year_avg**
4. Checked **Labels**
5. Left all other settings as default
6. Clicked **OK**

Excel generated a new sheet showing:

- **Rank** — position of each salary value relative to all others  
- **Percentile** — where each salary sits within the distribution (0–1 scale)

This makes it easy to identify:

- Highest‑paid roles  
- Lowest‑paid roles  
- Median and quartile positions  
- How individual salaries compare to the entire dataset  

Rank and percentile are especially useful when analysing salary competitiveness across job titles.

---

## 4. Moving Average (Understanding Peaks, Troughs & Trend Smoothing)

The moving average helps reveal the underlying trend in job postings by smoothing out daily fluctuations.  
This is important because job postings naturally have **peaks** (weekdays) and **troughs** (weekends), which can make the raw line chart noisy.

### Step 1 — Build a PivotTable for Daily Job Counts

1. Inserted a new PivotTable on the **jobs** table  
2. Added **job_posted_date** to **Rows**  
3. Extracted the **day** component so each row represents a single date  
4. Added **Count of job_posted_date** to **Values**  
5. Renamed the values field to **Job Count**

### Step 2 — Visualise the Raw Trend

- Insert → **PivotChart** → Line Chart  
- This shows the daily job posting pattern, including sharp drops on weekends

### Step 3 — Apply Moving Average Using Analysis ToolPak

Instead of relying on the PivotChart alone, I used the Analysis ToolPak for a cleaner moving average:

1. Deleted the PivotChart  
2. Data → **Data Analysis** → **Moving Average**
3. Input Range → the **Job Count** column
4. Interval → **7** (one full week)
5. Output Range → next column (replacing the “B” reference with the correct column)
6. Checked:
   - **Chart Output**
   - **Standard Error**

### Step 4 — Formatting

- Removed the legend  
- Cleaned up the x‑axis to make dates readable  
- Adjusted line thickness and colours for clarity  

### Why a 7‑day Moving Average?

- Captures one full weekly cycle  
- Smooths out weekday peaks and weekend troughs  
- Reveals the true hiring trend over time  
- Makes long‑term patterns easier to interpret  

This technique is essential for time‑series analysis, especially when the data has strong weekly seasonality like job postings.

---

## Summary

Today I learned how to:

- Use **Descriptive Statistics** to instantly generate summary metrics for salary data  
- Build **custom histograms** using the Analysis ToolPak for clearer visualisation  
- Apply **Ranking** and **Percentile** to compare salaries statistically  
- Use **Moving Averages** to smooth time‑series data and reveal trends  

These Analysis ToolPak tools automate complex calculations and make exploratory data analysis faster, more accurate, and more insightful.

# 📘 Day 31 — Introduction to Power Query (ETL in Excel)

## Overview

Today I started learning **Power Query**, a tool built into Excel that helps automate data cleaning and transformation. It is designed for real‑world messy data and allows me to build repeatable, reliable processes without writing code.

Power Query is essentially a visual **ETL tool**:

- **Extract** — bring data in from files, folders, databases, websites  
- **Transform** — clean, filter, merge, reshape, fix formats  
- **Load** — send the cleaned data back into Excel or the Data Model  

This makes it ideal for working with large or inconsistent datasets.

---

## What Power Query Does

### **1. Automates the ETL Process**
Instead of cleaning data manually every time, Power Query records each step once and replays it automatically when I click **Refresh**.  
This removes repetitive work and reduces errors.

### **2. Ensures Reproducibility**
Every transformation step appears in the **Applied Steps** panel.  
This means the process is:

- consistent  
- repeatable  
- easy to audit  
- easy for someone else to follow  

It works like a saved recipe.

### **3. Works Beyond Excel’s 1 Million Row Limit**
Excel sheets can only hold ~1,048,576 rows.  
Power Query processes data **in the background** (Data Model), so I can work with millions of rows even if the sheet itself cannot display them.

---

## Real‑World Analogy

If five staff members send daily sales reports in different formats, normally I would clean them manually.  
With Power Query, I set up the cleaning steps once — then click **Refresh** and everything updates automatically.

---

## Power Query vs Manual Work

| Without Power Query | With Power Query |
|---|---|
| Manual copy‑paste | One‑click refresh |
| Easy to make mistakes | Consistent steps |
| Hard to repeat | Fully reproducible |
| Needs coding for automation | Point‑and‑click interface |

---

## Getting Started in Excel

**Data tab → Get Data → From File → From Excel Workbook**

This opens the Power Query Navigator, where I can choose which sheet or table to load.

The interface shows:

- **Data preview** (centre)  
- **Applied Steps** (right)  
- **Queries list** (left)  

Each action I take becomes a recorded step.

---

## Importing a Single Excel File (First ETL Process)

To practise the basics, I imported one Excel file into another workbook:

1. Data → **Get Data**  
2. **From File → From Excel Workbook**  
3. Browsed to the file location  
4. Selected the file → **Import**  
5. Navigator window opened  
6. Selected the sheet I needed  
7. Clicked **Load**

This completed a full ETL cycle:

- Extracted the file  
- Transformed (if needed)  
- Loaded into Excel  

I also explored the **Queries & Connections** pane to see how Power Query stores and manages imported data.

---

## Summary

Today I learned:

- What Power Query is and why it exists  
- How it automates ETL and ensures reproducibility  
- How it handles datasets larger than Excel’s row limit  
- How to import a single Excel file using Power Query  
- How the interface works (Navigator, Applied Steps, Queries pane)

This sets the foundation for more advanced Power Query transformations in the next sessions.

