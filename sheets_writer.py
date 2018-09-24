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

    # Extract and print all of the values
    #list_of_hashes = sheet.get_all_values()
    last_row = 0
    for num in range(5, 100):
        temp_cell = sheet.acell('A' + str(num))
        if(temp_cell.value == ''):
            last_row = num
            break

    test_row = sheet_collector.gather_data(exported_data)
    for item in range(0, len(test_row)):
            client.login()
            sheet.update_cell(last_row, item+1, test_row[item])
