# UPS-Check
Pulls status from a list of APC UPS devices and emails an HTML Report.

<#==============================================================================
         File Name : UPS-Check.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Uses SNMP to poll and track APC UPS devices using MS Excel.
                   : 
             Notes : This script is INTENDED to use an Excel back-end to track UPS inventory.  Because 
                   : incidents arrise where Excel may not be available or function autonomously options 
                   : have been added to pull datas from both a flat text file as a CSV text file.
                   : Normal operation is with no command line options.  Commandline options noted below.
                   :
      Requirements : Requires the PowerShell SNMP library from https://www.powershellgallery.com/packages/SNMPv3
                   : Currently designed to poll APC UPS devices and emails a report.  UPS NMC must have SNMP v3 active.
                   : Script checks for active ping and SNMPv3.  Default operation is to check for a local text file 
                   : first, then if not found check for an existing Excel spreadsheet in the same folder or specified
                   : in the external config file.  If an existing spreadsheet is located the target list is compliled
                   : from column A.  Up to 10 copies of the HTML report are retained in a reports folder.  External
                   : config file example is at the end of the script.
                   : 
          Warnings : None
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   :   not limited to the following:
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.00 - 08-16-22 - Original 
    Change History : v2.00 - 09-00-22 - Numerous operational & bug fixes prior to v3.
                   : v3.10 - 09-15-22 - Cleaned up final version for posting.
                   : v4.00 - 04-12-23 - Too many changes to list
                   : v4.10 - 07-03-23 - Added age and LDOS dates. 
                   : v5.00 - 01-17-24 - Fixed DNS lookup.  Fixed last test result.  Fixed color coding of hostname for
                   :                    numerous events.  Added hostname cell comments to describe color coding.
                   : v6.00 - 02-12-24 - Retooled Html email report.  Added self test failed counts.  Added saved reports.
                   : v6.10 - 02-13-24 - Added missing external config entries.
                   : v7.00 - 02-16-24 - Fixed major bugs after moving config to external XML.
                   : v7.10 - 02-27-24 - Added exclusion list
                   : v7.20 - 03-05-24 - Fixed bugs found after PC crash.  Altered email sending options.
                   : v7.30 - 03-25-24 - Removed unknown status for everything that doesnt return that status from SNMP
                   : v7.40 - 12-24-24 - Fixed a number of typos.  Fixed detection of excluded IP addresses.
                   : v7.50 - 05-19-25 - Adjusted the way email sends.  Added try-catch for email and log file.     
                   : v8.00 - 06-12-25 - Shuffled Excel sections and added error checks in the event Excel automation
                   :                    fails due to PC restrictions when run as a scheduled task.  If Excel fails 
                   :                    the script will abort and produce an error log in the script folder.
                   : v8.10 - 08-01-25 - Fixed minor typos.  Updated module load routine.
                   : v8.20 - 09-03-25 - Edited SMNP import function.  Minor report text color adjustments.
                   : v8.30 - 09-05-25 - Rearranged columns in report for clarity.
                   : v8.40 - 02-06-26 - Added name of system running the script to the email report.  Adjusted when the 
                   :                    report header changes denoting failures and unknown status.
                   : v8.41 - 02-06-26 - Adjusted error in the way email send logic was formatted.
                   :                   
==============================================================================#>
