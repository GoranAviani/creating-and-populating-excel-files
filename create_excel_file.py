import openpyxl


def create_new_file():
    # Call a Workbook() function of openpyxl
    # to create a new blank Workbook object
    workbook = openpyxl.Workbook()

    # Get workbook active sheet
    # from the active attribute.
    sheet = workbook.active

    sheet_title = "business basic info"
    sheet.title = sheet_title

    print("active sheet title: " + sheet_title)

    # Note: The first row or column integer
    # is 1, not 0. Cell object is created by
    # using sheet object's cell() method.
    cell1 = sheet.cell(row=1, column=1)

    # writing values to cells
    cell1.value = "First name"

    c2 = sheet.cell(row=1, column=2)
    c2.value = "Last name"

    # Once have a Worksheet object, one can
    # access a cell object by its name also.
    # A2 means column = 1 & row = 2.
    c3 = sheet['A2']
    c3.value = "Goran"

    # B2 means column = 2 & row = 2.
    c4 = sheet['B2']
    c4.value = "Aviani"

    # Anytime you modify the Workbook object
    # or its sheets and cells, the spreadsheet
    # file will not be saved until you call
    # the save() workbook method.
    workbook.save("fist_example.xlsx")