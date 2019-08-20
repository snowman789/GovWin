import csv
import xlrd
import datetime

class Opportunity:
    def __init__(self, headers, row):
        self.opp_dict = {}
        self.hyper_link = ''
        index = 0
        for header in headers:
            self.opp_dict.update( {header : row[index] })
            index += 1
        self.mark = False
        self.delete = False
        try:
            str_lst = str(self.opp_dict['Solicitation Date']).split('/')
            if len(str_lst) == 2:
                self.solicitation_date = datetime.datetime(int(str_lst[1]),int(str_lst[0]), 1)
            if len(str_lst) == 3:
                self.solicitation_date = datetime.datetime(int(str_lst[2]), int(str_lst[0]), int(str_lst[1]) )
        except:
            print("Error: Award date info not found or unreadable. Setting award date to be in 01/01/3000")
            self.solicitation_date = datetime.datetime(1,1,3000)

        try:
            self.value = int(self.opp_dict['Value ($K)'])
        except:
            print("ERROR: Value was not found for opportunity, value of zero assigned instead")
            self.value = int(0)


headers = []


def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)

    opportunities = []

    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    # print(first_sheet.get_rows())

    mainData_book = xlrd.open_workbook(path, formatting_info=True)
    mainData_sheet = mainData_book.sheet_by_index(0)
    num_col = mainData_sheet.row_len(0)
    num_row = mainData_sheet.nrows
    # print(num_row)

    for row_index in range(0, num_row):
        if row_index == 0:
            for col in mainData_sheet.row(row_index):
                headers.append(col.value)
            # print("index of last updated: " + str(headers.index('Last Updated')))
        else:
            link = mainData_sheet.hyperlink_map.get((row_index, 0))
            url = '(No URL)' if link is None else link.url_or_path
            data_members = []
            for col in mainData_sheet.row(row_index):
                data_members.append(col.value)
            temp_opportunity = Opportunity(headers,data_members)
            temp_opportunity.hyper_link = url
            opportunities.append(temp_opportunity)

    print(opportunities[2].opp_dict['Last Updated'])
    print(opportunities[2].opp_dict['Projected Award Date'])
    # print(opportunities[2].opp_dict['Opp ID'])
    return opportunities

# ----------------------------------------------------------------------
if __name__ == "__main__":
    path = r'C:\Users\iroberts\Desktop\GovWin\test_excel_file.xls'
    open_file(path)
