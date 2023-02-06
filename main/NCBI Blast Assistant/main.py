import json
import xlsxwriter

# You Can Choose Your File Name Here:

workbook = xlsxwriter.Workbook('Result.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0


valid_data = {}

# Enter Your JSON File Directory Here:

with open('sample.json') as json_file:
    estDATA = json.load(json_file)

    for report in estDATA["BlastOutput2"]:

        hits = report['report']['results']['search']['hits']
        acc_num = report['report']['results']['search']['query_id']


        def hit_len_check():
            if len(hits) >= 1:
                proTitle = report['report']['results']['search']['hits'][0]['description'][0]['title']
                valid_data[acc_num] = proTitle
            elif len(hits) == 0:
                valid_data[acc_num] = "Null"

        hit_len_check()

# Transferring Data to the Spreadsheet
for key in valid_data.keys():
    
    worksheet.write(row, col,     key)
    worksheet.write(row, col + 1, valid_data[key])

    row += 1

workbook.close()