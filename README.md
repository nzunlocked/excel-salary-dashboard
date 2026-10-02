# Excel Salary Dashboard

## Overview

An interactive Microsoft Excel dashboard for analyzing data-related job
salaries across job titles, countries, and job schedule types.

The dashboard allows users to explore salary patterns and compare median
salaries across different roles and locations.

## Business Questions

- Which data-related job titles have the highest median salaries?
- How do median salaries vary across countries?
- How does job schedule type relate to salary?
- How do salary patterns differ between job roles and locations?

## Dataset

The dataset contains 2023 data-related job information, including:

- Job titles
- Salaries
- Countries
- Job schedule types
- Skills

## Tools & Skills

- Microsoft Excel
- Data Analysis
- Data Visualization
- Excel Formulas & Functions
- Dynamic Array Formulas
- Data Validation
- Dashboard Design

## Dashboard

![Salary Dashboard](0_Resources/Images/1_Salary_Dashboard_Final_Dashboard.gif)

## Key Analysis

### Salary by Job Title

A horizontal bar chart was used to compare median salaries across
different data-related job titles.

### Salary by Country

An Excel Map Chart was used to visualize median salaries across
countries represented in the dataset.

## Technical Implementation

### Median Salary Calculation

```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
