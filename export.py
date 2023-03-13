import openpyxl


def get_value(type, data):
    for infos in data:
        if infos[0] == type:
            return infos[1]
    return '/'


def to_excel(data_list):
    workbook = openpyxl.Workbook()

    worksheet = workbook.active

    label_values = {}
    list_company = []
    for company in data_list:
        result = {
            'Name': company[0][1],
            'Phone': get_value('Phone', company),
            'Website': get_value('Website', company),
            'Mail': get_value('Mail', company),
            'Employees': get_value('Effectifs', company),
            'Direction': get_value('Direction', company),
        }
        list_company.append(result)
    header_row = ['Name', 'Phone', 'Website', 'Mail', 'Employees', 'Direction']
    worksheet.append(header_row)
    for c in list_company:
        for i in range(0, len(c["Direction"])):
            if i == 0:
                row = [c['Name'], c['Phone'], c['Website'], c['Mail'], c['Employees'], c["Direction"][i]]
                worksheet.append(row)
            else:
                row = [None, None, None, None, None, c["Direction"][i]]
                worksheet.append(row)
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

    workbook.save("./output/data.xlsx")
