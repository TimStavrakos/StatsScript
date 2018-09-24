import os
import sheets_writer


spreadsheets_directory_string = '.\\CS Stat Spreadsheets'
spreadsheets_directory = os.fsencode(spreadsheets_directory_string)

#Open each spreadsheet to be written to the google sheet
for file in os.listdir(spreadsheets_directory):
    filename = os.fsdecode(file)
    if filename.endswith('xls') or filename.endswith('.xlsx'):
        sheets_writer.update_stats_sheet(spreadsheets_directory_string + '\\' + filename)

        #after the data from the current sheet has been uploaded, delete it
        os.remove(spreadsheets_directory_string + '\\' + filename)
