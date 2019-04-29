#!/usr/bin/env python3

#try to import pandas, else display error and close out of program.
try:
    import pandas
except:
    print('You need to install `Pandas` package to run this program. Run `pip install pandas` and try again, please!')
    exit()
#Sys is needed to input what file to parse.
import sys
#Pprint is a good tool for printing dictionaries.
import pprint
#used to open local html (and Rick Roll).
import webbrowser
#great for rounding to decimal places.
from decimal import Decimal

#function to convert from csv to Pandas.
def csv_to_pandas():
    #If there is a sys argument assign to variable.
    if len(sys.argv) > 1:
        vf_str_filename = sys.argv[1]
    #try to import csv from sys argument filename.
    try:
        vf_pand_main = pandas.read_csv(vm_str_filename,usecols=['AP Name','Operational Status','Client Count'])
    #if this fails, print error and then try default file ('WAPS1.csv').
    except:
        print('Inputed file not found, trying default file')
        print('')
        try:
           vf_pand_main = pandas.read_csv('WAPS1.csv',usecols=['AP Name','Operational Status','Client Count'])
        #If that fails, print error and exit program.
        except:
            print('Inputted and default files not found, please run program again and select correct file/directory!')
            exit()
    #if successful, return dataframe back to main.
    return(vf_pand_main)

#function to parse out information from the main pandas df.
def get_registration_info(vf_pand_main):

    #base challenge: create series with number of registered vs non-registered APs and print.
    vf_pand_registered = vf_pand_main.groupby('Operational Status').size()
    print(vf_pand_registered['Registered'],'registered APs')
    print(vf_pand_registered['Not Registered'],'not registered APs')
    print('')

    #Created df on only non-registered APs and export to excel
    vf_pand_not_reg = vf_pand_main.loc[vf_pand_main['Operational Status'] == 'Not Registered']
    try:
        vf_pand_not_reg.to_excel('notregistered.xlsx')
    except:
        print('I have found for Linux that openpyxl might need to be installed with `pip install openpyxl`')
    print('File notregistered.xlsx has been saved in local directory.  Enjoy!')
    print('---------------------')
    
    #make panda df of only registered APs
    vf_pand_reg = vf_pand_main.loc[vf_pand_main['Operational Status'] == 'Registered']

    #create pandas df of all rows with `Client Count` equal to zero.  Then use this to calculate the number of APs with this value and then sum of clients.
    #Repeat this for `Client Count` values of 1, 2, and 3.
    vf_pand_0client = vf_pand_reg[vf_pand_reg['Client Count'] == 0]
    vf_int_0_count = (vf_pand_0client.count())['Client Count']
    vf_int_0_sum = (vf_pand_0client.sum())['Client Count']

    vf_pand_1client = vf_pand_reg[vf_pand_reg['Client Count'] == 1]
    vf_int_1_count = (vf_pand_1client.count())['Client Count']
    vf_int_1_sum = (vf_pand_1client.sum())['Client Count']

    vf_pand_2client = vf_pand_reg[vf_pand_reg['Client Count'] == 2]
    vf_int_2_count = (vf_pand_2client.count())['Client Count']
    vf_int_2_sum = (vf_pand_2client.sum())['Client Count']
    
    vf_pand_3client = vf_pand_reg[vf_pand_reg['Client Count'] == 3]
    vf_int_3_count = (vf_pand_3client.count())['Client Count']
    vf_int_3_sum = (vf_pand_3client.sum())['Client Count']

    #Same process as above but for ranges. Range1 = 4-10 client, Range2 = 11-20 client, Range3 = 21-29 client, Hella = 30 or more clients.
    vf_pand_range1 = vf_pand_reg[(vf_pand_reg['Client Count'] >= 4) & (vf_pand_reg['Client Count'] <= 10)]
    vf_int_range1_count = (vf_pand_range1.count())['Client Count']
    vf_int_range1_sum = (vf_pand_range1.sum())['Client Count']

    vf_pand_range2 = vf_pand_reg[(vf_pand_reg['Client Count'] >= 11) & (vf_pand_reg['Client Count'] <= 20)]
    vf_int_range2_count = (vf_pand_range2.count())['Client Count']
    vf_int_range2_sum = (vf_pand_range2.sum())['Client Count']

    vf_pand_range3 = vf_pand_reg[(vf_pand_reg['Client Count'] >= 21) & (vf_pand_reg['Client Count'] <= 29)]
    vf_int_range3_count = (vf_pand_range3.count())['Client Count']
    vf_int_range3_sum = (vf_pand_range3.sum())['Client Count']
    
    vf_pand_hella = vf_pand_reg[(vf_pand_reg['Client Count'] >= 30)]
    vf_int_hella_count = (vf_pand_hella.count())['Client Count']
    vf_int_hella_sum = (vf_pand_hella.sum())['Client Count']

    #Calculate total number of registered APs plus total number of clients.
    vf_int_total_AP = vf_int_0_count + vf_int_1_count + vf_int_2_count + vf_int_3_count + vf_int_range1_count + vf_int_range2_count + vf_int_range3_count + vf_int_hella_count
    vf_int_total_client = vf_int_hella_sum + vf_int_range3_sum + vf_int_range2_sum + vf_int_range1_sum + vf_int_3_sum + vf_int_2_sum + vf_int_1_sum
    
    #Creat list of percentages, seems better because of multiple calculations per object. round(Decimal()) is used to have numbers rounded to two decimal places)
    vf_list_round_AP_perc = [round(Decimal(vf_int_0_count/vf_int_total_AP*100),2), round(Decimal(vf_int_1_count/vf_int_total_AP*100),2), round(Decimal(vf_int_2_count/vf_int_total_AP*100),2), round(Decimal(vf_int_3_count/vf_int_total_AP*100),2), round(Decimal(vf_int_range1_count/vf_int_total_AP*100),2), round(Decimal(vf_int_range2_count/vf_int_total_AP*100),2), round(Decimal(vf_int_range3_count/vf_int_total_AP*100),2), round(Decimal(vf_int_hella_count/vf_int_total_AP*100),2)]
    vf_list_round_client_perc = [0, round(Decimal(vf_int_1_sum/vf_int_total_client*100),2), round(Decimal(vf_int_2_sum/vf_int_total_client*100),2), round(Decimal(vf_int_3_sum/vf_int_total_client*100),2), round(Decimal(vf_int_range1_sum/vf_int_total_client*100),2), round(Decimal(vf_int_range2_sum/vf_int_total_client*100),2), round(Decimal(vf_int_range3_sum/vf_int_total_client*100),2), round(Decimal(vf_int_hella_sum/vf_int_total_client*100),2)]

    #Put all this information into a final results dataframe and pass back to main.
    vf_panda_results = pandas.DataFrame({'Clients': [0,1,2,3,'4-10','11-20','21-29','>=30'],
                                         '#APs': [vf_int_0_count, vf_int_1_count, vf_int_2_count, vf_int_3_count,vf_int_range1_count, vf_int_range2_count, vf_int_range3_count, vf_int_hella_count],
                                         '%Total APs': vf_list_round_AP_perc,
                                         '#Clients': [vf_int_0_sum, vf_int_1_sum, vf_int_2_sum, vf_int_3_sum, vf_int_range1_sum, vf_int_range2_sum, vf_int_range3_sum, vf_int_hella_sum],
                                         '%Total Clients': vf_list_round_client_perc})
    return(vf_panda_results)

#Menu selection to print report in Excel/HTML or both.
def choose_quest():
    print('Menu - Please select one of the following (or `q` to quit):')
    print('''
    1. Print Report to EXCEL
    2. Print Report to HTML
    3. Print to Both
          ''')     
    vf_str_select = input('Selection: ')
    #While Loop to retry is incorrect value is entered.
    while not(vf_str_select in ['1','2','3']):
        if vf_str_select == 'q':
            print('BYE FOR NOW!')
            exit()
        vf_str_select = input('Please select a valid number above or `q` to quit: ')
    return(int(vf_str_select))

#Function used if print to Excel option was selected.      
def pandas_to_excel(vf_panda_results):
    #Used tryp/except for xlsxwriter because it is only used for 2/3 menu selection options.
    try:
        import xlsxwriter
        #Write to 'pandas.xlsx' using xlsxwriter.
        writer = pandas.ExcelWriter('pandas.xlsx', engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object in Sheet1. Index=False to remove index
        vf_panda_results.to_excel(writer, sheet_name='Sheet1', index=False)
        # Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        #Create format object and apply to particular rows.
        format1 = workbook.add_format({'bg_color':   '#FFC7CE','font_color': '#9C0006'})
        worksheet.conditional_format('A2:E2', {'type': 'no_blanks','format':format1})
        worksheet.conditional_format('A3:E3', {'type': 'no_blanks','format':format1})
        worksheet.conditional_format('A9:E9', {'type': 'no_blanks','format':format1})
        writer.save()
        print('File pandas.xlsx has been saved in local directory.  Enjoy!')
    except:
        print('You need to install `Xlsxwriter` package to export to Excel. Run `pip install xlsxwriter` and try again, please!')

#Function to print into HTML and to open the HTML document in a browser.    
def pandas_to_HTML(vf_panda_results):
    #First try to use Jinja2 template to render HTML using formatting.
    #If that doesn't work (Jinja template missing or Jinja2 not installed) go from Pandas to HTML (no format).
    try:
        #import jinja2 (had it down here because other functions aren't dependent)
        from jinja2 import Template
        #Open template file and save to string.
        with open('jinjatemp.html','r') as myfile:
            vf_str_template = myfile.read()    
        tm = Template(vf_str_template)

        #Render HTML from template using a variety of variables from Pandas results df.
        vf_str_html = tm.render(
            vf_int_0_count = vf_panda_results['#APs'][0],
            vf_int_0_perc_AP = vf_panda_results['%Total APs'][0],
            vf_int_1_count = vf_panda_results['#APs'][1],
            vf_int_1_perc_AP = vf_panda_results['%Total APs'][1],
            vf_int_1_sum = vf_panda_results['#Clients'][1],
            vf_int_1_perc_client = vf_panda_results['%Total Clients'][1],
            vf_int_2_count = vf_panda_results['#APs'][2],
            vf_int_2_perc_AP = vf_panda_results['%Total APs'][2],
            vf_int_2_sum = vf_panda_results['#Clients'][2],
            vf_int_2_perc_client = vf_panda_results['%Total Clients'][2],
            vf_int_3_count = vf_panda_results['#APs'][3],
            vf_int_3_perc_AP = vf_panda_results['%Total APs'][3],
            vf_int_3_sum = vf_panda_results['#Clients'][3],
            vf_int_3_perc_client = vf_panda_results['%Total Clients'][3],
            vf_int_r1_count = vf_panda_results['#APs'][4],
            vf_int_r1_perc_AP = vf_panda_results['%Total APs'][4],
            vf_int_r1_sum = vf_panda_results['#Clients'][4],
            vf_int_r1_perc_client = vf_panda_results['%Total Clients'][4],
            vf_int_r2_count = vf_panda_results['#APs'][5],
            vf_int_r2_perc_AP = vf_panda_results['%Total APs'][5],
            vf_int_r2_sum = vf_panda_results['#Clients'][5],
            vf_int_r2_perc_client = vf_panda_results['%Total Clients'][5],
            vf_int_r3_count = vf_panda_results['#APs'][6],
            vf_int_r3_perc_AP = vf_panda_results['%Total APs'][6],
            vf_int_r3_sum = vf_panda_results['#Clients'][6],
            vf_int_r3_perc_client = vf_panda_results['%Total Clients'][6],
            vf_int_hella_count = vf_panda_results['#APs'][7],
            vf_int_hella_perc_AP = vf_panda_results['%Total APs'][7],
            vf_int_hella_sum = vf_panda_results['#Clients'][7],
            vf_int_hella_perc_client = vf_panda_results['%Total Clients'][7])
        #Save HTML to file.
        file = open('filename.html','w')
        file.write(vf_str_html)
        file.close()

    #If failure, should still be able to export generic HTML via Pandas to_html.
    except:
        print('')
        print('Error exporting HTML with formatting, exporting without formatting. You might need to install jinja2 with `pip install Jinja2`!')
        vf_panda_results.to_html('filename.html',index=False)
    #Either way, open in webbrowser.
    webbrowser.open('filename.html')
    
#Main program.       
if __name__ == '__main__':
    #First import csv to pandas dataframe.
    vm_pand_main = csv_to_pandas()
    #Parse out parts to be used in report.
    vm_panda_results = get_registration_info(vm_pand_main)
    #Menu selection function to determine output options.
    vm_int_select = choose_quest()
    if vm_int_select == 1:
        pandas_to_excel(vm_panda_results)
    elif vm_int_select == 2:
        pandas_to_HTML(vm_panda_results)
    elif vm_int_select == 3:
        pandas_to_excel(vm_panda_results)
        pandas_to_HTML(vm_panda_results)
