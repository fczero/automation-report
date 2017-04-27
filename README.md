# automation-report
CCP Daily Automation Report

 The report will be saved as an XLSX in the same directory.

 1.      Login to the VPN
 2.a     Add executable permission using chmod
         Ex. $chmod 775 ccp_daily_automation.py
 3       Run using $./ccp_daily_automation.py
   OR
 2.b     Run using $python3 ccp_daily_automation.py
    Help
     usage: ccp_daily_automation.py [-h] [-s | -r | -t]

     Scrape Jenkins report and create XLSX report.

     optional arguments:
         -h, --help        show this help message and exit
         -s, --smoke       Generate Smoke Report
         -r, --regression  Generate Regression Report
         -t, --test        Test mode, Smoke Confirmation
