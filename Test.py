import xlrd


def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)



    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    print(first_sheet.get_rows())

    mainData_book = xlrd.open_workbook(path, formatting_info=True)
    mainData_sheet = mainData_book.sheet_by_index(0)
    for row in range(1, 5):


        link = mainData_sheet.hyperlink_map.get((row, 0))
        url = '(No URL)' if link is None else link.url_or_path
        print( url)



# ----------------------------------------------------------------------
if __name__ == "__main__":

    path = r'C:\Users\iroberts\Desktop\GovWin\test_excel_file.xls'
    open_file(path)
