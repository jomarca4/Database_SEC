import xlwings as xw 
import pandas as pd 
import sqlite3

def read_queries():
    conn = sqlite3.connect('/Users/josemanuel/Desktop/financial_statements_SEC_EDGAR.db', timeout=10.0, factory=sqlite3.Connection)
    cursor = conn.cursor()

    query= '''
   SELECT DISTINCT companies.name, taxonomy.label, quarters.quarter_number, financial_statement_items.value, financial_statement_items.unit_of_measurement, financial_statements.date
FROM financial_statement_items
JOIN financial_statements ON financial_statement_items.financial_statement_id = financial_statements.id
JOIN quarters ON financial_statements.quarter_id = quarters.id
JOIN companies ON quarters.company_id = companies.id
JOIN taxonomy ON taxonomy.name = financial_statement_items.account_label
WHERE financial_statement_items.account_label LIKE '%even%' AND quarters."year"  =2019 AND financial_statements."type" = 'income_statement'
ORDER BY financial_statement_items.value DESC;
   
    '''
    query= '''
       SELECT * from companies
       '''
    # Define the variables
    company_names = 'Apple%'
    year = 2022
    #quarter_number = '1'

    # Construct the query with parameter placeholders
    query = '''
SELECT DISTINCT companies.name, taxonomy.name,taxonomy.label, quarters.quarter_number, financial_statement_items.value, financial_statement_items.unit_of_measurement, financial_statements.date
FROM financial_statement_items
JOIN financial_statements ON financial_statement_items.financial_statement_id = financial_statements.id
JOIN quarters ON financial_statements.quarter_id = quarters.id
JOIN companies ON quarters.company_id = companies.id
JOIN taxonomy ON taxonomy.name = financial_statement_items.account_label
WHERE quarters."year"=2022  AND (companies.name = 'WALMART INC.' OR companies.name LIKE '%APPLE%') AND quarters.quarter_number  ='2' AND financial_statements."type" = 'income_statement'
ORDER BY financial_statement_items.value DESC;
    '''

        # Execute the query and fetch the results
    
    wb = xw.Book("/Users/josemanuel/Desktop/Python_Scripts/income_stat_analysis.xlsx")
    sheet = wb.sheets["Sheet1"]
    sheet.clear()
    revenues = pd.read_sql_query(query, conn)
    sheet.range("A1").value = revenues

    
#read_queries()

def count_records_per_quarter():
    conn = sqlite3.connect('/Users/josemanuel/Desktop/financial_statements_SEC_EDGAR.db', timeout=10.0, factory=sqlite3.Connection)
    cursor = conn.cursor()

    query= '''
       SELECT DISTINCT quarters.quarter_number, quarters.year, companies.name, financial_statements.type, financial_statement_items.account_label
       FROM companies
       JOIN quarters ON quarters.company_id = companies.id
       JOIN financial_statements ON financial_statements.id = quarters.id
       JOIN financial_statement_items ON financial_statement_items.id = financial_statements.id
       WHERE companies.name LIKE '%Apple Inc.%'
       ORDER BY quarters.year
       '''
    revenues = pd.read_sql_query(query, conn)
    print(revenues)

    print(revenues.groupby(['year','quarter_number','type']).count())
    wb = xw.Book("/Users/josemanuel/Desktop/Python_Scripts/income_stat_analysis.xlsx")
    sheet = wb.sheets["Sheet1"]
    sheet.clear()
    revenues = pd.read_sql_query(query, conn)
    sheet.range("A1").value = revenues
count_records_per_quarter()

def count_records():
    conn = sqlite3.connect('/Users/josemanuel/Desktop/financial_statements_SEC_EDGAR.db', timeout=10.0, factory=sqlite3.Connection)
    cursor = conn.cursor()

    query= '''
SELECT quarters.quarter_number, quarters.year, COUNT(*) as num_records
FROM companies
JOIN quarters ON quarters.company_id = companies.id
       JOIN financial_statements ON financial_statements.id = quarters.id
       JOIN financial_statement_items ON financial_statement_items.id = financial_statements.id
       WHERE companies.name LIKE '%%'
GROUP BY quarters.year, quarters.quarter_number
ORDER BY quarters.year, quarters.quarter_number;
       '''
    revenues = pd.read_sql_query(query, conn)

    print(revenues)
count_records()