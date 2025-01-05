import pandas as pd
from datetime import datetime
import numpy as np

# Create Excel writer
writer = pd.ExcelWriter('xxxxxxx.xlsx', engine='xlsxwriter')

# Generate dates for the year
Dates = pd.date_range(start='2025-01-01', end='2025-12-31', freq='D')
num_days = len(Dates)

# Dashboard DataFrame
df_dashboard = pd.DataFrame({
    'KPI': ['Total Projects', 'In Progress', 'Done', 'Due'],
    'Value': [0, 0, 0, 0],
    'Meta': [0, 0, 0, 0]
})

# Resource Management DataFrame
df_resources = pd.DataFrame({
    'ID_Engineer': ['AAA', 'BBB', 'CCC', 'DDD', 'EEE', 'HHH'],
    'Name': ['GUY1', 'GUY2', 'GUY3', 'GUY4', 'GUY6', 'GUY7'],
    'Country': ['FR', 'BG', 'DE', 'DE', 'NL', 'NL'],
    'Competencie': ['Mechanic', 'Electric', 'Programming', 'Validation', 'Support', 'Support'],
    'Availability': [75, 100, 100, 100, 100, 100]
})

# Planning DataFrame
df_planning = pd.DataFrame({
    'ID': [''] * num_days,
    'Week': [0] * num_days,
    'Date': Dates,
    'Machine': [''] * num_days,
    'Serial': [''] * num_days,
    'Location': [''] * num_days,
    'Status': [''] * num_days,
    'Type': [''] * num_days,
    'Employee': [''] * num_days,
    'Activity': [''] * num_days,
    'Comments': [''] * num_days
})

# Add weekend identification
df_planning['IsWeekend'] = df_planning['Date'].dt.dayofweek.isin([5, 6])

# Action Log DataFrame
num_actions = len(df_resources)
df_action = pd.DataFrame({
    'Number': range(1, num_actions + 1),
    'Date': pd.to_datetime([datetime.now()] * num_actions),
    'Originator': [''] * num_actions,
    'Component': [''] * num_actions,
    'Description': [''] * num_actions,
    'Assigned': [''] * num_actions,
    'Due_Date': pd.to_datetime([datetime.now()] * num_actions),
    'Status': [''] * num_actions
})

# Write DataFrames to Excel
df_dashboard.to_excel(writer, sheet_name='Dashboard', index=False)
df_planning.to_excel(writer, sheet_name='Planning', index=False)
df_resources.to_excel(writer, sheet_name='Resources', index=False)
df_action.to_excel(writer, sheet_name='Action_Log', index=False)

# Get workbook and worksheet objects
workbook = writer.book
worksheet_dashboard = writer.sheets['Dashboard']
worksheet_planning = writer.sheets['Planning']
worksheet_resources = writer.sheets['Resources']
worksheet_action = writer.sheets['Action_Log']

# Add dashboard formulas
worksheet_dashboard.write_formula('B2', '=COUNTA(Planning!A:A)-1')
worksheet_dashboard.write_formula('B3', '=COUNTIF(Planning!G:G,"In Progress")')
worksheet_dashboard.write_formula('B4', '=COUNTIF(Planning!G:G,"Done")')
worksheet_dashboard.write_formula('B5', '=COUNTIF(Planning!G:G,"Due")')

# Format settings
formato_weekend = workbook.add_format({
    'bg_color': '#E6E6E6',
    'font_color': '#808080'
})
formato_porcentaje = workbook.add_format({'num_format': '0%'})

# Apply validations
worksheet_planning.data_validation('G2:G1000', {
    'validate': 'list',
    'source': ['Open', 'In Progress', 'Done', 'Due'],
    'input_title': 'Status',
    'input_message': 'Select the current status'
})

worksheet_planning.data_validation('H2:H1000', {
    'validate': 'list',
    'source': ['PRO', 'R&D', 'MAINT', 'SALES', 'INT'],
    'input_title': 'Type',
    'input_message': 'Select activity type'
})

worksheet_planning.data_validation('D2:D1000', {
    'validate': 'list',
    'source': ['VVVV', 'VVVV10', 'UUUU', 'UUUU10'],
    'input_title': 'Machine',
    'input_message': 'Select machine type'
})

worksheet_planning.data_validation('I2:I1000', {
    'validate': 'list',
    'source': '=Resources!$A$2:$A$7',
    'input_title': 'Select Employee',
    'input_message': 'Please select from the list of available employees'
})

# Apply weekend formatting
worksheet_planning.conditional_format('A2:K1000', {
    'type': 'formula',
    'criteria': '=WEEKDAY(C2,2)>5',
    'format': formato_weekend
})

# Format resources
worksheet_resources.set_column('E:E', None, formato_porcentaje)

# Adjust column widths
for worksheet in writer.sheets.values():
    worksheet.set_column('A:Z', 15)

# Save and close
writer.close()
