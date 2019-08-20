import configparser
import PullCSVData
import Create_Excel
import datetime
import os
def Manipulate_Data(data_path, save_path):
    # __location__ = os.path.realpath(
    #     os.path.join(os.getcwd(), os.path.dirname(__file__)))
    configFilePath = os.getcwd() + '\config.txt'
    # configFilePath = os.path.join(__location__, 'config.txt')
    # print(configFilePath)
    opportunities_to_pass = []
    configParser= configparser.RawConfigParser()

    configParser.read(configFilePath)

    mark_lst = []
    delete_lst = []

    config = configParser['DEFAULT']
    sections = configParser.sections()

    temp = config['Report Columns'].split(',')
    headers = []
    for phrase in temp:
        headers.append(str(phrase).strip())
    for section in sections:
        if configParser[section]['Mark phrases'] != '':
            mark_lst.append(section)
        if configParser[section]['Delete phrases'] != '':
            delete_lst.append(section)


    opportunities = PullCSVData.open_file(data_path)

    for opportunity in opportunities:
        for header in delete_lst:
            try:
                opp_phrase = str(opportunity.opp_dict[header]).lower()
                for phrase in str(configParser[header]['Delete phrases']).split(','):
                    phrase = str(phrase).lower()
                    if phrase in opp_phrase and phrase != '':
                        opportunity.delete = True


            except:
                print("There was a problem searching header: '" + str(header) + "' please check"
                                                                                " your spellig")


        if(not opportunity.delete):
            for header in mark_lst:
                try:
                    opp_phrase = str(opportunity.opp_dict[header]).lower()


                    for phrase in str(configParser[header]['Mark phrases']).split(','):
                        phrase = phrase.lower()
                        # print(phrase)

                        if phrase in  opp_phrase and phrase != '':
                            opportunity.mark = True
                except:
                    print("There was a problem searching header: '" + str(header) + "' please check"
                                                                               " your spelling")



        if not opportunity.delete:
            opportunities_to_pass.append(opportunity)

    if(config['Sort by Solicitation Date'].lower() == 'true'):
        try:
            opportunities_to_pass.sort(key = lambda opportunity: opportunity.solicitation_date)
        except:
            print("ERROR: Opportunities could not be sorted by date")

    # print(str(config['Sort by Value ($K)'].lower)+ "   HALP")
    if(config['Sort by Value'].lower() == 'true'):
        try:
            opportunities_to_pass.sort(key = lambda opportunity: int(opportunity.value))
        except:
            print("ERROR: Opportunities could not be sorted by Value")


    Create_Excel.Create_Excel_File(save_path, headers, opportunities_to_pass)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    data_path = r'C:\Users\iroberts\Desktop\GovWinV2\test_excel_file.xls'
    save_path = r'C:\Users\iroberts\Desktop\GovWinV2\test_results_12.xls.xlsx'
    Manipulate_Data(data_path, save_path)

# tst_lst = config['testlst']
# print(tst_lst.split(','))

# for item in config:
#
#     print(config[item])