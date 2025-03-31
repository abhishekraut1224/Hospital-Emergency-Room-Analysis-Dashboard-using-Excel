
![image](https://github.com/user-attachments/assets/05345e43-4f14-40c5-b420-fae2c19c5889)

# Hospital Emergency Room Analysis Dashboard

## Project Title
### Hospital Emergency Room Analysis Dashboard in Excel

## Objective / Purpose
We need to create a Hospital Emergency Room Analysis Dashboard in Excel to improve efficiency and provide useful insights. This dashboard will help stakeholders monitor, analyze, and make better decisions for managing patients and improving services.

## Database Overview
The dataset includes patient details such as age, gender, wait time, and department referrals. The data was cleaned and formatted to ensure consistency before creating the dashboard.

## Steps and Approach
#### - Business Requirement Gathering - Understanding stakeholder needs and defining dashboard objectives.

#### - Understanding the Data - Reviewing the dataset structure and identifying relevant fields.

#### - Data Connection - Importing data into Excel from various sources.

#### - Data Cleaning (Power Query) - Removed duplicates, handled missing values, and ensured correct formatting.

#### Creating the Calendar Table

### = List.Dates(#date(2023,01,01),731,#duration(1,0,0,0))

##### Reason: The calendar table is added to facilitate time-based analysis and build relationships with the main data table.

#### - Data Modeling - Building relationships between the calendar table (date column) and the main data table.

#### -Creating KPIs and Calculated Columns as per stakeholder requirements:

#### - Age Group Classification (DAX Formula):

### =IF([Patient Age]>=70,"70-79",IF([Patient Age]>=60,"60-69",IF([Patient Age]>=45,"45-59",IF([Patient Age]>=30,"30-44",IF([Patient Age]>=15,"15-29",IF([Patient Age]>=5,"05-14","0-4")))))))

#### Patient Attend Status:

### = IF ([Patient Waittime] < 30, "Within Time", "Delay")

#### - Creating Pivot Tables with respect to the charts and KPIs.

#### - Creating Slicers
- Year Slicer
- Month Slicer

### Adding a Clear Filters Button
Assigned a macro to reset all slicers and pivot table filters.
#### Macro Code:
Sub ClearAllSlicers()
    Dim slc As SlicerCache
    
    ' Loop through all slicers in the workbook and clear them
    For Each slc In ActiveWorkbook.SlicerCaches
        slc.ClearManualFilter
    Next slc
    
    ' Optional: Clear PivotTable Filters
    Dim pt As PivotTable
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.ClearAllFilters
        Next pt
    Next ws
    
    MsgBox "All slicers and filters have been cleared!", vbInformation, "Clear Filters"
End Sub

#### Acknowledgment: I took the help of ChatGPT for this macro.

## Visualization and Insights
The dashboard includes the following visualizations
Pie Charts: 
Patient attended on time
Number of patients by gender

Horizontal Bar Graphs:
Number of patients by departmental referral
Number of patients by age group

### KPI’s Requirement
The dashboard focuses on these key performance indicators (KPIs):

#### - Number of Patients - Total count of patients in the dataset

#### - Average Wait Time - Average duration patients waited before being attended

#### - Patient Satisfaction Score - Overall rating based on patient feedback

## Solutions

Built an interactive dashboard in Excel that allows stakeholders to analyze patient flow, efficiency, and department trends.
Used Power Query for data cleaning and transformation.
Created a calendar table to support date-based analysis.
Built relationships between the dataset and calendar table for better insights.
Developed slicers and a macro-based filter reset button for easier navigation.
Provided key visualizations to support decision-making and hospital efficiency improvements.


# Conclusion:
The dashboard successfully provides a clear, visual representation of hospital performance and patient flow.

Departments experiencing high referrals can reallocate resources to reduce bottlenecks.

If long wait times are an issue, optimizing staffing schedules, triage efficiency, and patient handling procedures can improve service.

Gender and age distribution insights can influence targeted healthcare improvements for specific patient groups.


## 📬 Contact
For any queries or suggestions, feel free to connect:
- **LinkedIn**: [Abhishek Mahadev Raut](https://www.linkedin.com/in/abhishek-raut-215191249/)
- **GitHub**: [abhishekraut1224](https://github.com/abhishekraut1224)
- **Mail**: abhiraut1224@gmail.com

---
⭐ If you found this project useful, consider giving it a **star** on GitHub!
