# Philippine Labor Market Case Study

## The Impact of COVID-19 on Labor Force, Employment, and Unemployment (2016–2025)

---

## About the Project

This project examines the impact of the COVID-19 pandemic on the Philippine labor market from 2016 to 2025 using labor force, employment, and unemployment data.

The analysis looks at the labor market before, during, and after the pandemic, with a focus on:

- The impact of COVID-19 on employment and unemployment
- Differences in unemployment between men and women
- The recovery of the Philippine labor market after the pandemic
- Whether key labor market indicators returned to or surpassed pre-pandemic levels

The project demonstrates hands-on experience with **Excel, Power Query, Pivot Tables, Pivot Charts, and Slicers** for data cleaning, analysis, visualization, and dashboard development.

---

## Research Objectives

1. Assess the impact of the COVID-19 pandemic on the Philippine labor market using labor force, employment, and unemployment data from 2016–2025.
2. Determine which gender was more adversely affected in terms of unemployment during and after the pandemic.
3. Evaluate whether the Philippine labor market recovered to or surpassed its pre-pandemic condition.

---

## Key Findings

### 1. Impact of COVID-19

The COVID-19 pandemic caused a significant disruption to the Philippine labor market in 2020.

- Unemployment rate increased from approximately **5% before the pandemic** to **10.26% in 2020**.
- Employment rate decreased from **94.89% to 89.74%**.
- The number of unemployed people increased from approximately **2.26 million to 4.50 million**.
- Labor force participation also declined as some people stopped looking for work during the pandemic.

### 2. Gender Differences in Unemployment

The results varied depending on the period being examined.

From **2016 to 2020**, men had higher unemployment rates than women. In 2020, male unemployment peaked at **10.61%**, compared with **9.68% for women**.

Starting in **2021**, female unemployment became higher and remained above male unemployment through 2025.

By 2025:

| Gender | Unemployment Rate |
|---|---:|
| Female | 4.35% |
| Male | 4.07% |

Although the number of unemployed men was higher in raw counts throughout the period, this is partly because more men are in the labor force. Unemployment rate provides a fairer comparison because it measures unemployed people relative to each gender's own labor force.

### 3. Labor Market Recovery

The overall Philippine labor market recovered and eventually surpassed several pre-pandemic indicators.

- Unemployment rate fell below the **2019 baseline by 2023**.
- Employment rate surpassed its 2019 level around **2022–2023**.
- Labor force participation also surpassed its 2019 level around **2022–2023**.
- However, the recovery was not completely equal across genders, as female unemployment remained higher than male unemployment from 2021 onward.

---

## Data Source

**Philippine Statistics Authority (PSA) – OpenSTAT**

**Dataset:**  
Labor Force Survey (LFS), Levels and Rates of Key Employment Indicators by Sex, Annual, 2016–2025

**Source:**  
https://openstat.psa.gov.ph/PXWeb/pxweb/en/DB/DB__1B__LFS/?rxid=44458430-2737-45a1-a793-fa5a25af1f22&tablelist=true

---

## Data Cleaning and Preparation

All data cleaning was performed using **Excel Power Query**. The **Applied Steps** feature was used to keep the transformation process traceable and repeatable.

### Cleaning Steps

1. Loaded the raw PSA dataset into Power Query.
2. Removed title and description rows that were not part of the actual dataset.
3. Promoted the first row to column headers.
4. Filtered to **Annual** data only, removing monthly records.
5. Limited the analysis to **2016–2025**.
6. Removed unnecessary columns, including:
   - Underemployment
   - Month
7. Renamed columns using shorter, cleaner labels.
8. Changed columns to their appropriate data types.
9. Removed rows containing errors.
10. Added labor market rate calculations.
11. Added a **Pandemic Period** classification:
    - Pre-Pandemic: 2016–2019
    - Pandemic: 2020–2022
    - Post-Pandemic: 2023–2025
12. Added recovery index calculations using **2019 as the baseline (100)**.
13. Corrected the scale of population, labor force, employed, and unemployed count columns by multiplying them by **1,000**. The PSA source represented thousands using decimal values (for example, `34090.501` represented `34,090,501`).
14. Applied comma-separated number formatting to the corrected count columns for readability.

### Retained Columns

The cleaned dataset retained the following fields:

- Year
- Pandemic Period
- Total Population
- Total Population Male
- Total Population Female
- Total Labor Force
- Total Labor Force Male
- Total Labor Force Female
- Employed Total
- Employed Male Total
- Employed Female Total
- Unemployed Total
- Unemployed Male Total
- Unemployed Female Total

---

## Calculated Columns and Formulas

The following calculated columns were created to support the analysis.

| Column | Formula | Description |
|---|---|---|
| **Labor Force Participation Rate** | `=Total Labor Force / Total Population` | Proportion of the total population participating in the labor force |
| **Employment Rate** | `=Employed Total / Total Labor Force` | Proportion of the labor force that is employed |
| **Unemployment Rate** | `=Unemployed Total / Total Labor Force` | Proportion of the labor force that is unemployed |
| **Male Employment Rate** | `=Employed Male Total / Total Labor Force Male` | Employment rate isolated to male figures |
| **Female Employment Rate** | `=Employed Female Total / Total Labor Force Female` | Employment rate isolated to female figures |
| **Male Unemployment Rate** | `=Unemployed Male Total / Total Labor Force Male` | Male-specific unemployment rate |
| **Female Unemployment Rate** | `=Unemployed Female Total / Total Labor Force Female` | Female-specific unemployment rate |
| **Pandemic Period** | `=IF(Year<=2019,"Pre-Pandemic",IF(Year<=2022,"Pandemic","Post-Pandemic"))` | Classifies each year into a pandemic period |
| **Unemployment Rate Index** | `=Unemployment Rate / $Q$5 * 100` | Indexes each year's unemployment rate against the fixed 2019 baseline |
| **Employment Rate Index** | `=Employment Rate / $P$5 * 100` | Indexes each year's employment rate against the fixed 2019 baseline |
| **LFPR Index** | `=Labor Force Participation Rate / $O$5 * 100` | Indexes each year's participation rate against the fixed 2019 baseline |

The `$O$5`, `$P$5`, and `$Q$5` references represent the cells containing the fixed 2019 baseline values in the original Excel workbook.

---

## Recovery Index

To evaluate recovery, **2019 was selected as the pre-pandemic baseline and assigned an index value of 100**.

### Unemployment Rate Index

```excel
=Unemployment Rate / 2019 Unemployment Rate * 100
```

### Employment Rate Index

```excel
=Employment Rate / 2019 Employment Rate * 100
```

### Labor Force Participation Rate Index

```excel
=Labor Force Participation Rate / 2019 Labor Force Participation Rate * 100
```

### Index Interpretation

| Index Value | Meaning |
|---|---|
| **100** | Same level as the 2019 baseline |
| **Above 100** | Higher than the 2019 baseline |
| **Below 100** | Lower than the 2019 baseline |

---

## Formulas Used in Analysis

### Percentage-Point Difference: Female − Male

```excel
=U2-T2
```

This calculates the difference between the female and male unemployment rates.

- **Positive value:** Female unemployment rate is higher
- **Negative value:** Male unemployment rate is higher
- **Zero:** Both rates are equal

### Period-Average Gender Comparison

```excel
=AVERAGEIF(PandemicPeriodRange,"Pandemic",FemaleRateRange)
-AVERAGEIF(PandemicPeriodRange,"Pandemic",MaleRateRange)
```

This calculates the difference between the average female and male unemployment rates during the pandemic period.

`AVERAGEIF` was used to average only the values matching the selected period.

---

## Excel Functions, Features, and Tools Used

| Function / Feature | Where Used | Description |
|---|---|---|
| **Power Query – Applied Steps** | Data cleaning | Records each cleaning action in a traceable and repeatable sequence |
| **Data Types** | Cleaning / preparation | Ensures columns are stored as the correct type, such as whole number, decimal, or text |
| **Number Formatting** | Count columns | Adds comma separators to large values for readability |
| **IF / Nested IF** | Pandemic Period | Classifies each year as Pre-Pandemic, Pandemic, or Post-Pandemic |
| **Absolute Cell Reference (`$` / F4)** | Recovery Index | Locks the 2019 baseline cell so it does not shift when the formula is copied |
| **AVERAGEIF** | Gender comparison | Calculates averages based on a specified text criterion |
| **Pivot Table** | Recovery Index | Summarizes Unemployment Rate Index, Employment Rate Index, and LFPR Index by year |
| **Pivot Chart** | Recovery Index visualization | Visualizes the recovery indexes as a line chart linked to the Pivot Table |
| **Slicer** | Interactive dashboard | Allows users to filter by Pre-Pandemic, Pandemic, Post-Pandemic, or year |
| **Report Connections** | Dashboard | Allows a slicer to control multiple connected Pivot Charts simultaneously |

---

## Pivot Table Analysis

A Pivot Table was created for the **Recovery Index** analysis.

### Configuration

- **Rows:** Year
- **Values:**
  - Unemployment Rate Index
  - Employment Rate Index
  - LFPR Index
- **Report Filter:** Pandemic Period
- **Aggregation:** Average

The **Average** aggregation was used because the values being summarized are rates and indexes rather than additive count values.

A **Pivot Chart** was then linked to the Pivot Table to visualize changes in the three recovery indexes over time.

---

## Interactive Dashboard

The dashboard uses **Slicers** to provide interactive filtering.

Users can select:

- Pre-Pandemic
- Pandemic
- Post-Pandemic
- Specific years

The slicers are connected to the relevant Pivot Charts through **Report Connections**, allowing the dashboard visuals to update simultaneously.

---

## Dashboard Components

The dashboard presents:

- Labor force trends
- Employment trends
- Unemployment trends
- Male vs. female unemployment
- COVID-19 impact
- Post-pandemic recovery
- Recovery indexes compared with 2019
- Interactive period/year filtering

---

## Tools Used

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Data analysis, formulas, Pivot Tables, and dashboard development |
| **Power Query** | Data cleaning and transformation |
| **Pivot Tables** | Data summarization |
| **Pivot Charts** | Data visualization |
| **Slicers** | Interactive filtering |
| **Excel Formulas** | Calculated indicators and analysis |

---

## Skills Demonstrated

- Data Cleaning
- Data Preparation
- Exploratory Data Analysis
- Excel
- Power Query
- Pivot Tables
- Pivot Charts
- Data Visualization
- Trend Analysis
- Gender Comparison
- Recovery Analysis
- Dashboard Development
- Working with Public/Government Datasets
- Translating raw data into meaningful insights

---

## Project Takeaway

The analysis shows that COVID-19 caused a sharp deterioration in the Philippine labor market in 2020, followed by a gradual recovery in the years that followed.

By 2023, major labor market indicators had recovered to or surpassed their 2019 levels. However, the gender comparison shows that the recovery was not completely uniform, with female unemployment remaining higher than male unemployment from 2021 onward.

Overall, this project demonstrates how a publicly available dataset can be cleaned, transformed, analyzed, and presented through an interactive dashboard to communicate meaningful trends and comparisons.

---

## Author

**Sherjames Compania**

B.S. Computer Science  
Major Focus on Data Science  
Central Philippine University

---

## Data Reference

**Philippine Statistics Authority (PSA) – OpenSTAT**

Labor Force Survey (LFS), Levels and Rates of Key Employment Indicators by Sex, Annual, 2016–2025.

https://openstat.psa.gov.ph/

---

## Disclaimer

This project is intended for **educational, analytical, and portfolio purposes**. The analysis and interpretations presented are my own and do not represent official conclusions of the Philippine Statistics Authority.
