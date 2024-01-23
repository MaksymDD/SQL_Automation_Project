# SQL Automation Project

## Project Description

Hello and welcome to my second project. This time, I'm combining SQL, VBA, and Python to automate the creation of Pivot Tables and charts in Excel.

In this project, I use Python, utilizing **Psycopg** as a database adapter, to connect to a SQL server. Using the **pandas** library, I extract data from our dataset to Excel. Within Excel, I leverage VBA to create charts and pivot tables.

## Technologies Used

- PostgreSQL
- pgAdmin
- Python
- VBA

For this project I am using e-commerce datasets avilible under link https://www.kaggle.com/datasets/carrie1/ecommerce-data. 

# Workflow

**PART1**

First, let me provide a snippet of Python code that connects to our dataset. 

```python
import psycopg2
import pandas as pd
import xlwings as xw

# PostgreSQL connection 
conn = psycopg2.connect(
    dbname="retails_transactions",
    user="postgres",
    password="kilovaTi2?",
    host="localhost",
    port="5432"
)
```

After that, I create SQL queries. For this project, I used three different queries. 

- **Total revenue per country**
- Total revenue per item
- Sales per quarter based from the year 2011
- 
However, for the purpose of demonstatrion I am going to use just one of them: revenue per country.


```python
#  SQL query per coutry
sql_query_country = """
    SELECT
        country,
        SUM(unit_price * quantity) AS total_revenue
    FROM
        transactions
    GROUP BY
        country
    ORDER BY
        total_revenue DESC;
"""

# Import data into a Pandas DataFrame
dfcountry = pd.read_sql_query(sql_query_country, conn)
```

And importing data into our Excel. For this purpose, I already have an Excel file created with appropriate worksheets.

```python
# Write data to an existing sheet 
sheet = wb.sheets["ByCountry"]
sheet.range("A1").value = dfcountry
"""
```

Execution of VBA

```python
wb.macro('RevenuePerCountry').run()
```

**PART 2**

VBA that create chart based on data from SQL. 

```vba
Sub RevenuePerCountry()
    Dim dataRange As Range
    Dim chartDestination As Range
    Dim chartObject As chartObject
    Dim chartSheetName As String
    Dim sheetCounter As Integer
    Dim tbl As ListObject
    Dim lastRow As Long


    ' data range for the bar chart
    Set dataRange = Sheets("ByCountry").Range("A1:B" & Cells(Rows.Count, "A").End(xlUp).Row)

    ' Sheet counter
    sheetCounter = 1
    
    ' Loop for name of new sheet
 Do
    chartSheetName = "Revenue per Country"
    If sheetCounter > 1 Then
        chartSheetName = chartSheetName & sheetCounter
    End If
    sheetCounter = sheetCounter + 1
Loop While WorksheetExists(chartSheetName)

    ' Create a new chart on the new sheet
    Set chartObject = Sheets.Add(After:=Sheets(Sheets.Count)).ChartObjects.Add(Left:=10, Width:=800, Top:=75, Height:=400)
    chartObject.Name = "LogRevenueBarChart"

    ' Set the data range for the bar chart
    chartObject.Chart.SetSourceData Source:=dataRange, PlotBy:=xlColumns

    ' Set the chart type to a bar chart
    chartObject.Chart.ChartType = xlBarClustered
    chartObject.Chart.BarShape = xlBox
    chartObject.Chart.Axes(xlValue).ScaleType = xlScaleLogarithmic

    ' Set the chart title
    chartObject.Chart.HasTitle = True
    chartObject.Chart.ChartTitle.Text = "Logarithmic Revenue Distribution by Country"

    ' Set the axis labels
    With chartObject.Chart.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Revenue"
    End With
```

And our function

```vba
Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
```

## Outputs

Here are results:

![data_per_country](https://github.com/MaksymDD/Duplicates_Project/assets/156451919/7fddbb58-97a9-430a-8b2e-b41791764b51)

*Data extracted from our SQL dataset*

![chart_per_country](https://github.com/MaksymDD/Duplicates_Project/assets/156451919/85a36ca9-76d8-4070-86e6-f06b20a43dbe)

*Chart of revenue per country*





    


