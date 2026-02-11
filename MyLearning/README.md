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










