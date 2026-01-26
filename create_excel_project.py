import pandas as pd
import numpy as np
import xlsxwriter

# Load the raw data we generated earlier
df = pd.read_csv('Manufacturing_Inventory_Raw_Data.csv')

# Create a new Excel workbook
workbook_path = 'Inventory_Planning_Project.xlsx'
writer = pd.ExcelWriter(workbook_path, engine='xlsxwriter')
workbook = writer.book

# 1. Raw_Data Sheet
df.to_excel(writer, sheet_name='Raw_Data', index=False)
worksheet_raw = writer.sheets['Raw_Data']

# 2. Inventory_Calculations Sheet
# We'll get unique SKUs and set up formulas
unique_skus = df[['SKU_ID', 'Category', 'Supplier_Name', 'Unit_Cost', 'Plant_Location']].drop_duplicates('SKU_ID')
unique_skus.to_excel(writer, sheet_name='Inventory_Calculations', index=False)
worksheet_calc = writer.sheets['Inventory_Calculations']

# Add headers for formulas
headers = ['Avg_Monthly_Demand', 'Demand_StdDev', 'CV (%)', 'Lead_Time_Days', 'Safety_Stock', 'Reorder_Point', 'Current_Stock', 'Status']
for col_num, header in enumerate(headers):
    worksheet_calc.write(0, 5 + col_num, header)

num_skus = len(unique_skus)
for i in range(1, num_skus + 1):
    sku_cell = f'A{i+1}'
    # Average Demand from Raw_Data (using SUMIF/COUNTIF for compatibility)
    worksheet_calc.write_formula(i, 5, f'=AVERAGEIF(Raw_Data!B:B, {sku_cell}, Raw_Data!H:H)')
    # StdDev helper (approximate for the demo as Excel STDEVIF is complex)
    worksheet_calc.write_formula(i, 6, f'=STDEV.P(IF(Raw_Data!$B$2:$B$1000={sku_cell}, Raw_Data!$H$2:$H$1000))', None, 0)
    # CV
    worksheet_calc.write_formula(i, 7, f'=IFERROR(G{i+1}/F{i+1}, 0)')
    # Lead Time (Lookup from Raw_Data)
    worksheet_calc.write_formula(i, 8, f'=XLOOKUP({sku_cell}, Raw_Data!B:B, Raw_Data!F:F)')
    # Safety Stock (Z=1.65 for 95%)
    worksheet_calc.write_formula(i, 9, f'=1.65 * G{i+1} * SQRT(I{i+1}/30)')
    # ROP
    worksheet_calc.write_formula(i, 10, f'=(F{i+1} * I{i+1}/30) + J{i+1}')
    # Current Stock (Last closing stock from Raw_Data)
    # For simplicity, we just lookup. In a real sheet, user would use XLOOKUP with match_mode
    worksheet_calc.write_formula(i, 11, f'=XLOOKUP({sku_cell}, Raw_Data!B:B, Raw_Data!J:J, 0, 0, -1)')
    # Status
    worksheet_calc.write_formula(i, 12, f'=IF(L{i+1}<=K{i+1}, "REORDER", "OK")')

# 3. ABC_XYZ_Analysis Sheet
analysis_df = unique_skus[['SKU_ID']].copy()
analysis_df.to_excel(writer, sheet_name='ABC_XYZ_Analysis', index=False)
worksheet_abc = writer.sheets['ABC_XYZ_Analysis']

abc_headers = ['Annual_Value', 'Value_Rank', 'ABC_Class', 'CV_Class', 'Final_Category']
for col_num, header in enumerate(abc_headers):
    worksheet_abc.write(0, 1 + col_num, header)

for i in range(1, num_skus + 1):
    sku_cell = f'A{i+1}'
    # Annual Value = Avg Demand * 12 * Unit Cost
    worksheet_abc.write_formula(i, 1, f'=Inventory_Calculations!F{i+1} * 12 * Inventory_Calculations!D{i+1}')
    # Rank
    worksheet_abc.write_formula(i, 2, f'=RANK(B{i+1}, B:B)')
    # ABC Class (Simple threshold for demo)
    worksheet_abc.write_formula(i, 3, f'=IF(C{i+1}<=(0.2*{num_skus}), "A", IF(C{i+1}<=(0.5*{num_skus}), "B", "C"))')
    # XYZ Class based on CV from previous sheet
    worksheet_abc.write_formula(i, 4, f'=IF(Inventory_Calculations!H{i+1}<0.2, "X", IF(Inventory_Calculations!H{i+1}<0.5, "Y", "Z"))')
    # Combined
    worksheet_abc.write_formula(i, 5, f'=D{i+1} & E{i+1}')

# 4. Supplier_Scorecard
suppliers = df['Supplier_Name'].unique()
supp_df = pd.DataFrame({'Supplier': suppliers})
supp_df.to_excel(writer, sheet_name='Supplier_Scorecard', index=False)
worksheet_supp = writer.sheets['Supplier_Scorecard']

supp_headers = ['Total_Orders', 'Delays', 'OTD (%)']
for col_num, header in enumerate(supp_headers):
    worksheet_supp.write(0, 1 + col_num, header)

for i in range(1, len(suppliers) + 1):
    supp_cell = f'A{i+1}'
    worksheet_supp.write_formula(i, 1, f'=COUNTIF(Raw_Data!E:E, {supp_cell})')
    worksheet_supp.write_formula(i, 2, f'=COUNTIFS(Raw_Data!E:E, {supp_cell}, Raw_Data!L:L, 1)')
    worksheet_supp.write_formula(i, 3, f'=1 - (C{i+1}/B{i+1})')

# 5. Scenario_Analysis
worksheet_scenario = workbook.add_worksheet('Scenario_Analysis')
worksheet_scenario.write('A1', 'Scenario Inputs', workbook.add_format({'bold': True}))
worksheet_scenario.write('A2', 'Demand Surge %')
worksheet_scenario.write('B2', 0.2, workbook.add_format({'num_format': '0%'}))
worksheet_scenario.write('A3', 'Lead Time Delay %')
worksheet_scenario.write('B3', 0.1, workbook.add_format({'num_format': '0%'}))

worksheet_scenario.write('A5', 'Impact Summary', workbook.add_format({'bold': True}))
worksheet_scenario.write('A6', 'Metric')
worksheet_scenario.write('B6', 'Baseline')
worksheet_scenario.write('C6', 'Simulated')

worksheet_scenario.write('A7', 'Avg Reorder Point')
worksheet_scenario.write('B7', '=AVERAGE(Inventory_Calculations!K:K)')
worksheet_scenario.write('C7', '=B7 * (1 + B2 + B3)')

# 6. Dashboard
worksheet_dash = workbook.add_worksheet('Dashboard')
worksheet_dash.write('B2', 'INVENTORY PERFORMANCE DASHBOARD', workbook.add_format({'bold': True, 'font_size': 20}))
worksheet_dash.write('B4', 'KPIs', workbook.add_format({'bold': True}))
worksheet_dash.write('B5', 'Total Inventory Value')
worksheet_dash.write('C5', '=SUM(ABC_XYZ_Analysis!B:B)')
worksheet_dash.write('B6', 'Critical SKUs (Stockout Risk)')
worksheet_dash.write('C6', '=COUNTIF(Inventory_Calculations!M:M, "REORDER")')
worksheet_dash.write('B7', 'Avg Supplier OTD%')
worksheet_dash.write('C7', '=AVERAGE(Supplier_Scorecard!D:D)')

# Close the writer
writer.close()
print(f'Excel workbook created: {workbook_path}')
