import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime, timedelta
import string


def get_client(json_file_name):
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
    return gspread.authorize(creds)


def get_time(now: datetime, delta):
    hour = now.hour
    minutes_datetime = now + (datetime.min - now) % delta
    if minutes_datetime.minute is 15:
        hour += 0.25
    elif minutes_datetime.minute is 30:
        hour += 0.5
    elif minutes_datetime.minute is 45:
        hour += 0.75
    return hour


def main():
    with open('config.json') as json_data_file:
        data = json.load(json_data_file)

    # get client with credentials
    client = get_client(data['credentials_path'])
    # get document by id
    document = client.open_by_key(data['spreadsheet_key'])

    # get sheet in worksheet
    worksheets = document.worksheets()
    current_month_string = datetime.now().strftime('%B')
    if len([x for x in worksheets if x.title == current_month_string]) is 0:
        raise ValueError("No sheet for the current month in spreadsheet!")
    sheet = document.worksheet(current_month_string)

    # get cell for current day
    first_data_cell = data['first_data_cell']
    first_data_cell[1] += datetime.now().day - 1
    column = string.ascii_lowercase.index(first_data_cell[0].lower()) + 1
    row = first_data_cell[1]
    if sheet.cell(col=column, row=row).value is not "":
        raise ValueError(
            "There is already some Date in Cell ({}, {})!".format(column, row))

    # update cell with data
    time = get_time(datetime.now(), timedelta(minutes=data['round_to_minutes']))

    sheet.update_cell(col=column, row=row, value=time)


if __name__ == '__main__':
    main()
