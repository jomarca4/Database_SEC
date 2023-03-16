import xlwings as xw 
import pandas as pd 
import sqlite3

def read_queries():
    conn = sqlite3.connect('/Users/josemanuel/Desktop/Python_Scripts/financial_statements_SEC_EDGAR.db')

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

    query = '''
SELECT DISTINCT companies.name, taxonomy.label, quarters.quarter_number, financial_statement_items.value, financial_statement_items.unit_of_measurement, financial_statements.date
FROM financial_statement_items
JOIN financial_statements ON financial_statement_items.financial_statement_id = financial_statements.id
JOIN quarters ON financial_statements.quarter_id = quarters.id
JOIN companies ON quarters.company_id = companies.id
JOIN taxonomy ON taxonomy.name = financial_statement_items.account_label
WHERE companies.name = 'Apple Inc.' AND quarters."year"  =2022 AND quarters.quarter_number  ='' AND financial_statements."type" = 'income_statement'
ORDER BY financial_statement_items.value DESC;





'''

        # Execute the query and fetch the results
    result = conn.execute(query).fetchall()
    revenues = pd.read_sql_query(query, conn)
    wb = xw.Book("/Users/josemanuel/Desktop/Python_Scripts/income_stat_analysis.xlsx")
    sheet = wb.sheets["Sheet1"]
    sheet.clear()
    sheet.range("A1").value = revenues

    
read_queries()
