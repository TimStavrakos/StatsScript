import os
import sheets_writer


spreadsheets_directory_string = 'F:\\Temp\\CS Stat Spreadsheets'
spreadsheets_directory = os.fsencode(spreadsheets_directory_string)

for file in os.listdir(spreadsheets_directory):
    filename = os.fsdecode(file)
    if filename.endswith('xls') or filename.endswith('.xlsx'):
        sheets_writer.update_stats_sheet(spreadsheets_directory_string + '\\' + filename)
        os.remove(spreadsheets_directory_string + '\\' + filename)
