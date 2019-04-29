# __Johnny's Wireless Access Point Report__

> This document uses Markdown.  Please view on Github (https://github.com/johnnywilkes/NPL_APR_2019_WAP) or use the following for viewing: https://stackedit.io/app#

## ___Overall Program Idea___

> **Please note that this program uses the [Pandas](http://pandas.pydata.org) module.  If you don't have the Pandas module installed, it will prompt and error and quit.  Run `pip install pandas` to install.  I have also found that Ubuntu requires openpyxl to be installed for the to_excel Pandas method.  Also, if the final report is selected to export to Excel, xlsxwriter will need to be installed as well.**

> **Jinja2 is also used in this program to render HTML (print report to HTML option).  However, because there is an option to export to HTML without formatting, I allowed that options as well as a fallback (if Jinja2 isn't installed).**

> **This program also assumes that you have selected the csv to parse via command line argument or that you are going to parse a default file ('WAPS1.csv') in the same directory as the script.**


The goal of this challenge was to take a csv file (AP report from Cisco Prime) and to extract data from in in multiple manners:
 - Find the number of unregistered  APs found,
 - Optionally list the unregistered  APs in a clear, concise and readable format.
 - Create a report which summarizes the number of clients per registered AP.  Format to highlight APs with low (zero or one) and high (more than 30) clients/AP.
 
I have accomplished all of these aspects of this challenge as well as giving the user some choice as far as printing the final report in Excel, HTML or both.

The generalized flow of my program is the following:

1. Import CSV to Pandas dataframe (csv_to_pandas).
2. Parse number of unregistered APs to print (get_registration_info).
3. Export dataframe of only unregistered APs to Excel document (get_registration_info).
4. Parse out information for final report (bonus) and store in python dictionary (get_registration_info).
5. Export python dictionary to Pandas dataframe (get_registration_info).
6. Give user a choice to print final report to Excel, HTML or both (choose_quest).
7. Export Pandas dataframe (or report info) to either Excel(pandas_to_excel), HTML(pandas_to_HTML) or both.


## ___Variable Naming/Program Structure___

This program uses the same variable naming, comment and program structure as last month's submission for, more information, see the section with the same name (Variable Naming/Program Structure) in the following link:
https://github.com/johnnywilkes/NPL_NOV_2018_TIME/blob/master/README.md


## ___Possible Refactoring/Feature Releases___

 - Seems like there were at least a few parts that could have been cleaned up/simplified if I had some more time and some peer review.  I think the sections with listing all variables to put into vf_panda_results dataframe as well as all the variables used in Jinja2 for HTML rendering could have been simplified somehow.
 - It would have been nice to have more time to help others with their projects. I did the best I could to encourage, motivate and coach others when possible.