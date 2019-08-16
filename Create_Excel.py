import xlsxwriter
import os
import PullCSVData
import Manipulate_Data

def GetAttributueList(opportunity, attr_lst):
    unique_lst = []

    for item in attr_lst:

        try:
            hold = opportunity.opp_dict[item]
            # print(hold)
            # hold = getattr(obj,item)
        except:
            hold = ''
        unique_lst.append(hold)
    return unique_lst

def Create_Excel_File(report_file_path, headers, opportunities):
    to_return = ''

    ## GENERATE OUTPUT
    # Create a workbook and add a worksheet.
    # workbook = xlsxwriter.Workbook(r'C:\Users\iroberts\Desktop\Billet_comparison_results\Billet_comparison.xlsx')
    # print(report_file_path)
    # report_file_path = r'C:\Users\iroberts\Documents\SR Reports\Generated_Report.xlsx'


    workbook = xlsxwriter.Workbook(report_file_path)
    worksheet = workbook.add_worksheet()
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, len(headers), 19)



    header_format = workbook.add_format()
    header_format.set_bold()
    header_format.set_bg_color('#add8e6')
    header_format.set_center_across()

    header_format.set_border()



    normal_format = workbook.add_format()
    normal_format.set_text_wrap()
    normal_format.set_border()

    needs_attention_format = workbook.add_format()
    needs_attention_format.set_text_wrap()
    needs_attention_format.set_bg_color('yellow')
    needs_attention_format.set_border()

    hyperlink_format = workbook.add_format()
    hyperlink_format.set_text_wrap()
    hyperlink_format.set_border()
    hyperlink_format.set_underline()
    hyperlink_format.set_font_color('blue')

    attn_hyperlink_format = workbook.add_format()
    attn_hyperlink_format.set_text_wrap()
    attn_hyperlink_format.set_border()
    attn_hyperlink_format.set_underline()
    attn_hyperlink_format.set_font_color('blue')
    attn_hyperlink_format.set_bg_color('yellow')

    try:
        temp = int(headers.index("Latest News"))
        worksheet.set_column(temp, temp, 40)
    except:
        x = 1

    name_index = int(headers.index("Program Name"))
    # worksheet.set_column('B:B', 50)
    # worksheet.set_column('K:K', 40)
    # worksheet.set_column('E:E', 12)

    row=0
    col=0

    item_number = 0
    # key_words = ['cyber', 'advance notice', 'advanced notice', 'not an rfp', 'rfi', 'sources sought', 'industry day']



    for item in headers:
        worksheet.write(row, col, item, header_format)
        col +=1
    col = 0
    row+=1

    mylst = []
    mylst.pop
    for opportunity in opportunities:
        attributes = GetAttributueList(opportunity,headers)
        format = normal_format


        if opportunity.mark:
            format = needs_attention_format

        for attribute in attributes:
            if col == name_index:
                format_link = hyperlink_format
                if opportunity.mark:
                    format_link = attn_hyperlink_format
                worksheet.write_url(row, col, opportunity.hyper_link, format_link, string=opportunity.opp_dict["Program Name"])
            else:
                worksheet.write(row, col, attribute, format)
            col += 1
        col = 0
        row += 1





    workbook.close()
    os.startfile(report_file_path)
    print('done')
    return 'Success!'

