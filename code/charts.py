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

# SQL queries
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

sql_query_item = """
    SELECT
        description,
        SUM(unit_price * quantity) AS total_revenue
    FROM
        transactions
    GROUP BY
        description
    ORDER BY
         total_revenue DESC
    LIMIT 25;
"""

sql_query_pivot = """
SELECT
    EXTRACT(YEAR FROM invoice_date) AS year,
    EXTRACT(QUARTER FROM invoice_date) AS quarter,
    country,
    SUM(unit_price * quantity) AS total_revenue
FROM
    transactions
WHERE
    EXTRACT(YEAR FROM invoice_date) = 2011
GROUP BY
    year, quarter, country
ORDER BY
    year, quarter, country;
"""

# Import data into a Pandas DataFrame
dfcountry = pd.read_sql_query(sql_query_country, conn)

dfitem = pd.read_sql_query(sql_query_item, conn)

dfpivot = pd.read_sql_query(sql_query_pivot, conn)

# Close the connection
conn.close()

# Excel file path - PARTH FOR THE XLSM FILE
excel_file_path = r"PLEASE PROVIDE YOUR PATH\test_template.xlsm"

# Opens the Excel file with xlwings
wb = xw.Book(excel_file_path)

# Write data to an existing sheet 
sheet = wb.sheets["ByCountry"]
sheet.range("A1").value = dfcountry

sheet = wb.sheets["Total_Rev"]
sheet.range("A1").value = dfitem

sheet = wb.sheets["Quarters"]
sheet.range("A1").value = dfpivot

wb.macro('RevenuePerItem').run()

wb.macro('RevenuePerCountry').run()

wb.macro('PivotQuarters').run()



