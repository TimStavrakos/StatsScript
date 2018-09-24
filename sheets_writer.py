import gspread
import stats_collector
from oauth2client.service_account import ServiceAccountCredentials
from gspread.httpsession import HTTPSession

def update_stats_sheet(exported_data):
    # use creds to create a client to interact with the Google Drive API
    headers = HTTPSession(headers={'Connection':'Keep-Alive'})
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.Client(creds, headers)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    client.login()
    sheet = client.open_by_key('1OspJrWAgzBj6Pm6TPio8O6ba6c9VQMTt3EjlAjl805Q').sheet1

    last_row = 0
    for num in range(5, 100):
        temp_cell = sheet.acell('A' + str(num))
        if(temp_cell.value == ''):
            last_row = num
            break

    new_row = stats_collector.gather_data(exported_data)

    '''We generate an entire row of data at once, but the connection times out
        when trying to add an entire row to the google sheet. In order to upload
        the row, we have to write each element of the array to it's corresponding
        cell in the sheet.'''
        
    for item in range(0, len(new_row)):
            client.login()
            sheet.update_cell(last_row, item+1, nwe_row[item])
