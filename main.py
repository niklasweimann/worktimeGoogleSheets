import string
from datetime import datetime, timedelta

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from config import Config


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
    elif minutes_datetime.minute is 0:
        hour += 1
    return hour


def map_to_minutes(param):
    minutes = 0
    if param is 0.25:
        minutes = 15
    elif param is 0.5:
        minutes = 30
    elif param is 0.75:
        minutes = 45
    return minutes


def updateCell(config: Config):
    # get client with credentials
    client = get_client(config.get_property("credentials_path"))
    # get document by id
    document = client.open_by_key(config.get_property("spreadsheet_key"))
    # get sheet in worksheet
    worksheets = document.worksheets()
    current_month_string = datetime.now().strftime('%B')
    if len([x for x in worksheets if x.title == current_month_string]) is 0:
        raise ValueError("No sheet for the current month in spreadsheet!")
    sheet = document.worksheet(current_month_string)

    # get cell for current day
    first_data_cell = config.get_property("first_data_cell")
    first_data_cell[1] += datetime.now().day - 1
    column = string.ascii_lowercase.index(first_data_cell[0].lower()) + 1
    row = first_data_cell[1]
    time = ""

    if sheet.cell(col=column, row=row).value is "":
        # update cell with start data
        if sheet.cell(col=column, row=row).value is not "":
            print("Der Arbeitsbeginn für heute wurde bereits eingetragen!")
            return
        print("Beginne Arbeitstag. Viel Erfolg :)")
        time = get_time(datetime.now(), timedelta(minutes=config.get_property("round_to_minutes")))
    elif sheet.cell(col=column+2, row=row).value is "":
        time = get_time(datetime.now(), timedelta(minutes=config.get_property("round_to_minutes")))
        column += 2
        if sheet.cell(col=column, row=row).value is "" and sheet.cell(col=column + 2, row=row).value is "":
            column += 2
            print("Beginne Pause. Guten Appetit :)")
        elif sheet.cell(col=column, row=row).value is "" and sheet.cell(col=column + 2, row=row).value is not "":
            break_begin = sheet.cell(col=column + 2, row=row).value.split(",")
            hours = int(break_begin[0])
            minutes = 0
            if len(break_begin) > 1:
                minutes = int(map_to_minutes(break_begin[1]))
            break_begin_time = datetime.now().replace(hour=hours, minute=minutes)
            time = time - get_time(break_begin_time, timedelta(minutes=config.get_property("round_to_minutes")))
            sheet.update_cell(col=column+2, row=row, value="")
            print("Beende Pause. Frohes Schaffen #workhardanywhere")
        else:
            print("Die Pause wurde für heute bereits eingetragen!")
            return

    elif sheet.cell(col=column+1, row=row).value is "":
        time = get_time(datetime.now(), timedelta(minutes=config.get_property("round_to_minutes")))
        column += 1
        if sheet.cell(col=column, row=row).value is not "":
            print("Der Arbeitsbeginn für heute wurde bereits eingetragen!")
            return
        print("Beende Arbeitstag. Schönen Feierabend!")

    sheet.update_cell(col=column, row=row, value=time)


def main():
    config = Config("config.json")
    if not config.is_valid():
        raise ValueError("Config-File is not valid")

    updateCell(config)



if __name__ == '__main__':
    main()
