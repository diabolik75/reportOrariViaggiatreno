
from datetime import datetime
from urllib.request import urlopen 

import json 
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
#hasmap numerotreno:codicestazione
treni = {
    "5838" : "S09211",
    "5834" : "S09211"
}
# current date and time
dataOdierna = datetime.now()
dataOdierna = dataOdierna.replace(hour=0, minute=0, second=0, microsecond=0)
#oggi = dataCompleta.date() 
print('Date and Time is:', dataOdierna)
timestamp = int(round(datetime.timestamp(dataOdierna)))
#print("timestamp =", timestamp)
workbook = xlsxwriter.Workbook('report_orari_treno_'+str(dataOdierna.year)+str(dataOdierna.month)+'.xlsx')
cell_format_standard = workbook.add_format()
cell_format_ritardo = workbook.add_format()
cell_date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
cell_stazioni_format = workbook.add_format()
cell_stazioni_format.set_rotation(45)
cell_stazioni_format.set_bg_color('silver')
for codiceTreno in treni:
    try:

        # store the URL in url as  
        # parameter for urlopen 
        url = "http://www.viaggiatreno.it/infomobilita/resteasy/viaggiatreno/andamentoTreno/"+treni[codiceTreno]+"/"+codiceTreno+"/"+str(timestamp)+"000"
        print("url =", url)
        # store the response of URL 
        response = urlopen(url) 
        # storing the JSON response  
        # from url in data 
        data_json = json.loads(response.read())   
        # print the json response 
        #print(data_json) 
        # Start from the cell . Rows and columns are zero indexed.
        row = dataOdierna.day
        col = 0
        orarioDiPartenza = datetime.fromtimestamp(int(round(data_json['orarioPartenzaZero']/1000)))
        #orarioDiPartenza = datetime.strptime(orarioDiPartenza, '%H:%M')
        #print(int(round(data_json['orarioPartenzaZero']/1000)))
        print("%s:%s" % (orarioDiPartenza.strftime("%H"), orarioDiPartenza.strftime("%M")))
        worksheet = workbook.add_worksheet(codiceTreno+'_'+data_json['origine'][0:3]+'_'+data_json['destinazione'][0:3]+'_'+orarioDiPartenza.strftime("%H")+orarioDiPartenza.strftime("%M"))
        worksheet.write(0,0, "Giorno")
        worksheet.write(33,0, "Totale Ritardo")
        worksheet.write(34,0, "Media Ritardo")
        worksheet.write(row,col, dataOdierna, cell_date_format)
        print("data odierna =", dataOdierna)
        # Iterating through the json
        # list
        col=1
        for fermate in data_json['fermate']:
            print(str(fermate['stazione']) + "," + str(fermate['ritardo']))
            worksheet.write(0,col, str(fermate['stazione']),cell_stazioni_format)
            if (fermate['ritardo']>0):
                cell_format_ritardo.set_bg_color('yellow')
                worksheet.write(row,col, fermate['ritardo'],cell_format_ritardo)
                print('yellow')
            else:
                cell_format_standard.set_bg_color(False)
                worksheet.write(row,col, fermate['ritardo'],cell_format_standard)
            column_letter= xlsxwriter.utility.xl_col_to_name(col)
            print (column_letter)
            worksheet.write(33,col, '=SUM(%s2:%s32)' %(column_letter,column_letter))
            worksheet.write(34,col, '=AVERAGE(%s2:%s32)' %(column_letter,column_letter))
            col = col+1
    except:
        print("Si è verificato un errore nel recupero delle informazioni")
workbook.close()   
