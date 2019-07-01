import xlrd
from dateutil import parser
from datetime import datetime


def get_events(path, sheet_index=0, key="SF"):
    """Given a path of an excel file, and the key word to search for,
    return a list of row name (part title) and column name (date)"""

    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(sheet_index)
    events = []

    # Scan the excel file for all cells that contanin the key ("SF") and return them
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            if (sheet.cell_value(i, j) == 'Date'):
                date_row = i
            if (sheet.cell_value(i, j) == key):
                events.append([sheet.cell_value(i, 0), str(parser.parse(sheet.cell_value(date_row, j)).date())])

    return events


def main():
    loc = ("E:\Programs\Script\Create Events From Excel\Sound-Platform-Attendants Complete List_March 2019-1 (1).xlsx")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            if (sheet.cell_value(i, j) == 'Date'):
                # Storing the row index for the date (since it is not 0)
                # so the event date can be found on that corresponding row
                date_row = i
            if (sheet.cell_value(i, j) == 'SF'):
                # Alternate solution
                # myDate = datetime.strptime(sheet.cell_value(date_row, j)+"-2019", '%d-%b-%Y')

                # Since the date is in 'dd-MMM' format, we need to convert it to isoformat
                print(sheet.cell_value(i, 0), str(parser.parse(sheet.cell_value(date_row, j)).date()))


if __name__ == '__main__':
    main()

# Program extracting all columns
# name in Python
# import xlrd
#
# loc = ("E:\Programs\Script\Sound-Platform-Attendants Complete List_March 2019-1 (1).xlsx")
#
# wb = xlrd.open_workbook(loc)
# sheet = wb.sheet_by_index(0)
#
# # For row 0 and column 0
# sheet.cell_value(0, 0)
#
# for i in range(sheet.ncols):
#     print(sheet.cell_value(0, i))
