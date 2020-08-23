import openpyxl


def create_new_file():
    # Call a Workbook() function of openpyxl
    # to create a new blank Workbook object
    workbook = openpyxl.Workbook()

    # Get workbook active sheet
    # from the active attribute.
    sheet = workbook.active

    # Once have the Worksheet object,
    # one can get its name from the
    # title attribute.
    sheet_title = "business basic info"
    sheet.title = sheet_title

    print("active sheet title: " + sheet_title)
