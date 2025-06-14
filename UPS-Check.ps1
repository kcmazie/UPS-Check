Param(
    [switch]$Console = $false,                                                  #--[ Set to true to enable local console result display. Defaults to false ]--
    [switch]$Debug = $False                                                     #--[ Generates extra console output for debugging.  Defaults to false ]--
)
<#==============================================================================
         File Name : UPS-Check.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Uses SNMP to poll and track APC UPS devices using MS Excel.
                   : 
             Notes : Normal operation is with no command line options.  Commandline options noted below.
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
                   : v7.51 - 05-29-25 - Added title color change if any anomaly is detected.
                   : v8.00 - 06-12-25 - Shuffled Excel sections and added error checks in the event Excel automation
                   :                    fails due to PC restrictions when run as a scheduled task.  If Excel fails 
                   :                    the script will abort and produce an error log in the script folder.
                   :                   
==============================================================================#>
#$ErrorActionPreference = "stop"
#try{
#Requires -version 5
Clear-Host 

#--[ Variables ]---------------------------------------------------------------
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Today = Get-Date -Format MM-dd-yyyy 
$CloseAnyOpenXL = $false

#--[ Runtime tweaks for testing ]--
$Console = $False
$Debug = $False
#------------------------------------------------------------

$ErrorActionPreference = "stop"
try{
    if (!(Get-Module -Name SNMPv3)) {
        Get-Module -ListAvailable SNMPv3 | Import-Module | Out-Null
    }
    Install-Module -Name SNMPv3
}Catch{
    Write-host "Error installing SNMP module" -ForegroundColor red
}

#==[ Functions ]===============================================================
Function SendEmail ($MessageBody,$ExtOption) { 
    $Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,$ExtOption.SmtpPort) 
    $Email = New-Object System.Net.Mail.MailMessage  
    $Email.IsBodyHTML = $true
    $Email.From = $ExtOption.EmailSender
    $Email.Subject = "UPS Status Report"
    $Email.Body = $MessageBody
    $ErrorActionPreference = "stop"
    Try {  #--[ If running out of an IDE console, in debug, or alwaysalt modes, send only to the alt user for testing ]-- 
        If (($ExtOption.ConsoleState -or $ExtOption.Debug) -or ($ExtOption.AlwaysAlt -like "*True*")){
            $Email.To.Add($ExtOption.EmailAltRecipient)  
            $Smtp.Send($Email)
            If ($ExtOption.ConsoleState){Write-Host `n"--- Email Sent ---" -ForegroundColor red }
        }Else{
            If ($ExtOption.Alert){  #--[ If a device failed self-test or trigger day is matched send to main recipient ]--     
                $Email.To.Add($ExtOption.EmailRecipient)          
                $Smtp.Send($Email)
                If ($ExtOption.ConsoleState){Write-Host `n"--- Email Sent ---" -ForegroundColor red }     
            }
        }
    }Catch{
        $ErrorMessage = $_.Error.Message
        $ExceptionMsg = $_.Exception.Message
        $Failed = $True
    }
    If ($ExtOption.Debug -or $Failed){
        If ($Failed){
            StatusMsg "-- Error sending email --" "Red" $ExtOption
        }
        $Msg="-- Email Parameters --" 
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Error Msg     = "+$_.Error.Message
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Exception Msg = "+$_.Exception.Message
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Local Sender  = "+$ThisUser
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Recipient     = "+$ExtOption.EmailRecipient
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="SMTP Server   = "+$ExtOption.SmtpServer
        StatusMsg $Msg "yellow" $ExtOption
    }
}

Function StatusMsg ($Msg, $Color, $ExtOption){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    If ($ExtOption.Debug){
        Add-Content -Path "$PSScriptRoot\error.log" -Value $Msg
    }
    Write-Host "-- Script Status: $Msg" -ForeGroundColor $Color
    $Msg = ""
}

Function LoadConfig ($ExtOption,$ConfigFile){  #--[ Read and load configuration file ]-------------------------------------
    StatusMsg "Loading external config file..." "Magenta" $ExtOption
    if (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
        [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--    
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SourcePath" -Value $Config.Settings.General.SourcePath
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelSourceFile" -Value $Config.Settings.General.ExcelSourceFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DNS" -Value $Config.Settings.General.DNS
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DayOfWeek" -Value $Config.Settings.General.DayOfWeek
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SNMPv3User" -Value $Config.Settings.Credentials.SNMPv3User
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SNMPv3AltUser" -Value $Config.Settings.Credentials.SNMPv3AltUser
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SNMPv3Secret" -Value $Config.Settings.Credentials.SNMPv3Secret
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PasswordFile" -Value $Config.Settings.Credentials.PasswordFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "KeyFile" -Value $Config.Settings.Credentials.KeyFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "IPlistFile" -Value $Config.Settings.General.IPListFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.General.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $Config.Settings.General.EmailRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailAltRecipient" -Value $Config.Settings.General.EmailAltRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "AlwaysAlt" -Value $Config.Settings.General.AlwaysAlt
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSender" -Value $Config.Settings.General.EmailSender
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "HNPattern" -Value $Config.Settings.General.HNPattern
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Exclusions" -Value $Config.Settings.Exclusions.Exclude
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $False
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $ExtOption
        break;break;break
    }
    If ((Get-Date).DayOfWeek -eq $ExtOption.DayOfWeek){  #--[ Triggers email to group on selected day of week ]--
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Alert" -Value $True
    }
    Return $ExtOption
}

Function GetConsoleHost ($ExtOption){  #--[ Detect if we are using a script editor or the console ]--
    Switch ($Host.Name){
        'consolehost'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $False -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell Console detected." -Force
        }
        'Windows PowerShell ISE Host'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell ISE editor detected." -Force
        }
        'PrimalScriptHostImplementation'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "COnsoleMessage" -Value "PrimalScript or PowerShell Studio editor detected." -Force
        }
        "Visual Studio Code Host" {
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "Visual Studio Code editor detected." -Force
        }
    }
    If ($ExtOption.ConsoleState){
        StatusMsg "Detected session running from an editor..." "Magenta" $ExtOption
    }
    Return $ExtOption
}

Function SNMPv3Walk ($Obj,$ExtOption,$OID){
    $WalkRequest = @{
        UserName   = $ExtOption.SNMPv3User
        Target     = $Obj.IPAddress
        OID        = $OID
        AuthType   = 'MD5'
        AuthSecret = $ExtOption.SNMPv3Secret
        PrivType   = 'DES'
        PrivSecret = $ExtOption.SNMPv3Secret
        #Context    = ''
    }
 $erroractionpreference = "stop"
    try{
        $Result = Invoke-SNMPv3Walk @WalkRequest | Format-Table -AutoSize
    }Catch{
        add-content -Path $PSScriptRoot\new-error.log -Value $_.Exception.Message
    }
    If ($ExtOption.Debug){write-host "SNMpv3 Debug :" $Result }
    Return $Result
}

Function GetSNMPv1 ($Obj,$ExtOption,$OID) {
    $SNMP = New-Object -ComObject olePrn.OleSNMP
    $erroractionpreference = "Stop"
    Try{
        $snmp.open($Obj.IPAddress,$ExtOption.SNMPv3User,2,1000)
        $Result = $snmp.get($OID)
    }Catch{
        $Result = $_.Exception.Message        
    }
    If ($ExtOption.Debug){ write-host "SNMpv1 Debug :" $Result }
    $Obj | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result -force
    Return $Obj
}

Function GetSNMPv3 ($Obj,$ExtOption,$OID){
   If ($Obj.v3User){  #--[ If set $true the main v3 user tested good so use it ]--
        $GetRequest1 = @{
            UserName   = $ExtOption.SNMPv3User
            Target     = $Obj.IPAddress
            OID        = $OID.Split(",")[1]
            AuthType   = 'MD5'
            AuthSecret = $ExtOption.SNMPv3Secret
            PrivType   = 'DES'
            PrivSecret = $ExtOption.SNMPv3Secret
        }

        Try{
            $Result = Invoke-SNMPv3Get @GetRequest1 #-ErrorAction:Stop
        }Catch{
            $Result = $_.Exception.Message
            If ($ExtOption.Debug){
                StatusMsg "SNMPv3 User 1 failed..." "red" $ExtOption
                StatusMsg $_.Exception.Message "cyan" $ExtOption
            }
        }
    }Else{  #--[ User 1 has failed so use user 2 instead ]--
        $GetRequest2 = @{
            UserName   = $ExtOption.SNMPv3AltUser
            Target     = $Target
            OID        = $OID.Split(",")[1]
                 #--[ Auth and Priv Not needed ]--
            #AuthType   = 'MD5'
            #AuthSecret = $Script:SNMPv3Secret
            #PrivType   = 'DES'
            #PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest2 -ErrorAction:Stop
        }Catch{
            $Result = $_.Exception.Message
            If ($ExtOption.Debug){
                StatusMsg "SNMPv3 User 2 failed..." "red" $ExtOption
                StatusMsg $_.Exception.Message "cyan" $ExtOption
            }
        }
    }

    If ($ExtOption.Debug){
        StatusMsg "  -- SNMPv3 Debug -- " 'Yellow' $ExtOption
        If ($Test){
            StatusMsg "SNMP User 2  " "Green" $ExtOption
        }Else{
            StatusMsg "SNMP user 1  " "Green" $ExtOption
        }
        StatusMsg $OID.Split(",")[0] "Cyan" $ExtOption
        StatusMsg $Result "yellow" $ExtOption
    }
    Return $Result
}

#--[ End of Functions ]-------------------------------------------------------

$OIDArray = @()
$OIDArray += ,@('LastTestResult','.1.3.6.1.4.1.318.1.1.1.7.2.3.0') #--[ 1=passed, 2=failed, 3=never run ]--
$OIDArray += ,@('LastTestDate','.1.3.6.1.4.1.318.1.1.1.7.2.4.0')  #--[ returns a date or nothing ]--
$OIDArray += ,@('UPSSerial','.1.3.6.1.4.1.318.1.1.1.1.2.3.0')
#$OIDArray += ,@('UPSModelName','.1.3.6.1.4.1.318.1.1.1.1.1.1.0')
$OIDArray += ,@('UPSModelName','.1.3.6.1.4.1.318.1.4.2.2.1.5.1')
$OIDArray += ,@('UPSModelNum','.1.3.6.1.4.1.318.1.1.1.1.2.5.0')
#$OIDArray += ,@('UPSModelNum2','1.3.6.1.4.1.318.1.1.1.1.1.1')
$OIDArray += ,@('UPSMfgDate','.1.3.6.1.4.1.318.1.1.1.1.2.2.0')
#--[ MfgDate from SN:  xx1915xxxxxx means mfg in 2019, 15th week.  ]--
#$OIDArray += ,@('UPSIDName','.1.3.6.1.2.1.33.1.1.5.0')
#$OIDArray += ,@('FirmwareVer','.1.3.6.1.4.1.318.1.1.1.1.2.1.0')
$OIDArray += ,@('Mfg','.1.3.6.1.2.1.33.1.1.1.0')
$OIDArray += ,@('MfgDate','.1.3.6.1.4.1.318.1.1.1.1.2.2.0')
#$OIDArray += ,@('MAC','.1.3.6.1.2.1.2.2.1.6.2')
$OIDArray += ,@('Location','.1.3.6.1.2.1.1.6.0')
#$OIDArray += ,@('Contact','.1.3.6.1.2.1.1.4.0')       
$OIDArray += ,@('HostName','.1.3.6.1.2.1.1.5.0')       
$OIDArray += ,@('NMC','.1.3.6.1.2.1.1.1.0')   
#$OIDArray += ,@('BattFreqOut','.1.3.6.1.4.1.318.1.1.1.4.2.2.0')
#$OIDArray += ,@('BattVOut','.1.3.6.1.4.1.318.1.1.1.4.2.1.0')
#3$OIDArray += ,@('BattVIn','.1.3.6.1.4.1.318.1.1.1.3.2.1.0')
#$OIDArray += ,@('BattFreqIn','.1.3.6.1.4.1.318.1.1.1.3.2.4.0')
#$OIDArray += ,@('BattActualV','.1.3.6.1.4.1.318.1.1.1.2.2.8.0')
#$OIDArray += ,@('BattCurrentAmps','.1.3.6.1.4.1.318.1.1.1.2.2.9.0')
#$OIDArray += ,@('BattChangedDate','.1.3.6.1.4.1.318.1.1.1.2.1.3.0')
#$OIDArray += ,@('BattCapLeft','.1.3.6.1.4.1.318.1.1.1.2.2.1.0')
$OIDArray += ,@('BattRunTime','.1.3.6.1.4.1.318.1.1.1.2.2.3.0')
#$OIDArray += ,@('BattReplace','.1.3.6.1.4.1.318.1.1.1.2.2.4.0')
#$OIDArray += ,@('BattReplaceDate','.1.3.6.1.4.1.318.1.1.1.2.2.21.0')
#$OIDArray += ,@('BattSKU','.1.3.6.1.4.1.318.1.1.1.2.2.19.0')
#$OIDArray += ,@('BattExtSKU','.1.3.6.1.4.1.318.1.1.1.2.2.20.0')
#$OIDArray += ,@('BattTemp','.1.3.6.1.4.1.318.1.1.1.2.2.2.0')
#$OIDArray += ,@('ACVIn','.1.3.6.1.4.1.318.1.1.1.3.2.1.0')
#$OIDArray += ,@('ACFreqIn','.1.3.6.1.4.1.318.1.1.1.3.2.4.0')
#$OIDArray += ,@('LastXfer','.1.3.6.1.4.1.318.1.1.1.3.2.5.0')
#$OIDArray += ,@('UPSVOut','.1.3.6.1.4.1.318.1.1.1.4.2.1.0')
#$OIDArray += ,@('UPSFreqOut','.1.3.6.1.4.1.318.1.1.1.4.2.2.0')
#$OIDArray += ,@('UPSOutLoad','.1.3.6.1.4.1.318.1.1.1.4.2.3.0')
#$OIDArray += ,@('UPSOutAmps','.1.3.6.1.4.1.318.1.1.1.4.2.4.0')    

#==[ Begin ]==============================================================

#--[ Load external XML options file ]--
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
$ExtOption = New-Object -TypeName psobject #--[ Object to hold runtime options ]--
$ExtOption = LoadConfig $ExtOption $ConfigFile

#--[ Detect Runspace ]--
$ExtOption = GetConsoleHost $ExtOption 
If ($ExtOption.ConsoleState){ 
    StatusMsg $ExtOption.ConsoleMessage "Cyan" $ExtOption
}
If ($Debug){
    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $True
}

StatusMsg "Processing UPS Devices" "Yellow" $ExtOption
$erroractionpreference = "stop"

If (Test-Path -Path $ExtOption.SourcePath -PathType leaf){
    $SourcePath = $ExtOption.SourcePath
}Else{
    $SourcePath = $PSScriptRoot
}

#--[ Locate and load data from source file ]--
#--[ If the text file exists the IP list will be pulled from it ]--
If (Test-Path -Path ($SourcePath+"\"+$ExtOption.IPListFile) -PathType Leaf){   
    $IPList = Get-Content ($SourcePath+"\"+$ExtOption.IPListFile)         
    StatusMsg "IP text list was found, loading IP list from it... " "green"  $ExtOption
}ElseIf (Test-Path -Path ($SourcePath+"\"+$ExtOption.ExcelSourceFile) -PathType Leaf){
    StatusMsg "IP text list not found, Attempting to process spreadsheet... " "cyan" $ExtOption
    #--[ Close copies of Excel that PowerShell has open ]--
    If ($CloseAnyOpenXL){
        $ProcID = Get-CimInstance Win32_Process | Where-Object {$_.name -like "*excel*"}
        ForEach ($ID in $ProcID){  #--[ Kill any open instances to avoid issues ]--
            Foreach ($Proc in (get-process -id $id.ProcessId)){
                if (($ID.CommandLine -like "*/automation -Embedding") -Or ($proc.MainWindowTitle -like "$ExcelWorkingCopy*")){
                    Stop-Process -ID $ID.ProcessId -Force
                    StatusMsg "Killing any existing open PowerShell instance of Excel..." "Red" $ExtOption
                    Start-Sleep -Milliseconds 100
                }
            }
        }
    }
    #--[ Create new Excel COM object ]--
    Try{
        $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        StatusMsg "Creating new Excel COM object..." "Magenta" $ExtOption
    }Catch{
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value "`r`n============================"
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value "Excel automation has failed @ $DateTime "
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value "Please save spreadsheet column A as a text file to process it."
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value "--[ Exception Msg ]---------"
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value $_.Exception.Message
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value "--[ Error Msg ]-------------"
        Add-Content -Path "$PSScriptRoot\--UPSCheck_FAILURE--.log" -Value $_.Error.Message
         break
    }
    $Excel.Visible = $false
    $WorkBook = $Excel.Workbooks.Open(($SourcePath+"\"+$ExtOption.ExcelSourceFile))
    $WorkSheet = $Workbook.WorkSheets.Item("UPS")
    $WorkSheet.activate()   
    $WorkSheet = $Workbook.WorkSheets.Item("UPS")
    $WorkSheet.activate()    
    $Row = 2   
    $IPList = @() 
    Do {
        $IPList += ,@($Row,$WorkSheet.Cells.Item($row,1).Text)
        $Row++
    } Until (
        $WorkSheet.Cells.Item($row,1).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
    )  
    $Excel.DisplayAlerts = $false
    $WorkBook.Close($true)
    $Excel.Quit()
    $Excel.Quit()
    $Excel.Quit()  
}Else{
    StatusMsg "Existing spreadsheet not found, Source copy failed, Nothing to process.  Exiting... " "red" $ExtOption
    Break;break
}

ForEach ($Target in $IPList){  #--[ Are we pulling from Excel or a text file?  Jagged has row numbers from Excel ]--
    if ($Target.length -eq 2){
        $Jagged = $True
    }Else{
        $Jagged = $False
    }
}

#==[ Process items collected in $IPList, from spreadsheet, or text file as appropriate ]===============================
$Row = 2   
$TestPass = 0
$TestFail = 0
$TestUnknown = 0
$Offline = 0
$Excluded = 0
$HtmlHeader = @() 
$HtmlReport = @() 
$HtmlBody = @()
$Count = $IPList.Count

ForEach ($Target in $IPList){
    $Obj = New-Object -TypeName psobject   #--[ Individual Target Device Result Collection ]--

    If ($Jagged){
        $Row = $Target[0]
        $Target = $Target[1]
    }
    $Current = $Row-3

    $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $Target -force
    
    If ($ExtOption.ConsoleState){
        Write-Host "`nCurrent Target  :"$Target"  ("$Current" of "$Count")" -ForegroundColor Yellow 
    }
   
    Try{
        $HostLookup = (nslookup $Obj.IPAddress $ExtOption.DNS 2>&1) 
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value ($HostLookup[3].split(":")[1].TrimStart()).Split(".")[0] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $True
    }Catch{
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value "Not Found" -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False
    }

    #--[ Read list of excluded IP's from XML ]--
    If ($ExtOption.Exclusions.Split(",") -contains $Obj.IPAddress){
        $Obj | Add-Member -MemberType NoteProperty -Name "Excluded" -Value $true -force
    }ElseIf (Test-Connection -ComputerName $Obj.IPAddress -count 1 -buffersize 16 -Quiet){  #--[ Ping target ]--
        $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Online" -force

        #--[ Test for SNMPv3 access.  Make sure to include leading comma on OID ]---------
        If ($ExtOption.ConsoleState){
            StatusMsg "Testing SNMPv3..." "Magenta" $ExtOption
        }

        If ((!($Debug)) -and ($ExtOption.ConsoleState)){
            Write-Host "  Working." -NoNewline
        }

        $Obj | Add-Member -MemberType NoteProperty -Name "v3User1" -Value $True  #--[ Test for valid v3 user ]--
        $Result = GetSNMPv3 $Obj $ExtOption ",1.3.6.1.2.1.1.8.0" 

        if ($Result -like "*TimeTicks*"){
            $Obj | Add-Member -MemberType NoteProperty -Name "v3User1" -Value $true -force
            $Obj | Add-Member -MemberType NoteProperty -Name "SNMP" -Value $true -force
        }Else{
            $Obj | Add-Member -MemberType NoteProperty -Name "v3User1" -Value $False -force
            $Obj | Add-Member -MemberType NoteProperty -Name "SNMP" -Value $True -force
        }

        If ((!($Debug)) -and ($ExtOption.Console)){
            Write-Host "." -NoNewline
        }

        #--[ Test for SNMPv1 if v3 user failed ]------------------------------------------
        If (!($Obj.SNMP)){
            $Result = GetSNMPv1 $Obj $ExtOption "1.3.6.1.2.1.1.8.0" 
            if ($Result -like "*TimeTicks*"){
                $Obj | Add-Member -MemberType NoteProperty -Name "SNMP" -Value $True -force
            }Else{
                $Obj | Add-Member -MemberType NoteProperty -Name "SNMP" -Value $False -force
            }
        }
    }

    #--[ Only process OIDs if online and SNMPv3 are both good ]--------------------------
    If (($Obj.Connection -eq "Online")){ #} -and ($Obj.SNMPv3)){  
        ForEach ($Item in $OIDArray){            
            $Result = GetSNMPv3 $Obj $ExtOption $Item 
            If ($Debug){
                $Msg = "DEBUG -- "+$Item[0]+'='+$Result
                StatusMsg $Msg "yellow" $ExtOption
            }Else{
                If ($ExtOption.ConsoleState){Write-Host "." -NoNewline}   #--[ Writes a string of dots to show operation is proceeding ]--
            }

            If ($Obj.HostName -like "*chill*" ){
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "Chiller" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "BattRunTime" -Value "N/A" -force
            }Else{
                $Obj | Add-Member -MemberType NoteProperty -Name "DeviceType" -Value "UPS" -force
            }

            #--[ Clean Up Results ]-------------------------------------------------
            $erroractionpreference = "silentlycontinue"
            Switch ($Item[0]) {
                "HostName" {   #--[ Extract and compare hostname ]--   
                    If ($Obj.Hostname -match $Result.Value.ToString()){
                        $SaveVal = $Result.Value.ToString()                  #--[ Hostnames match ]--
                    }Else{
                        If ($Obj.Hostname -like $ExtOption.HNPattern){
                            $SaveVal = $Obj.Hostname                         #--[ DNS value is good ]--    
                        }ELse{
                            $SaveVal = $Result.Value.ToString()              #--[ DNS wrong, use SNMP ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False -force
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "HostName" -Value $SaveVal -force
                } 
                "LastTestResult" {   #--[ Extract last test result ]--  
                    Switch ($Result.Value){
                        "1" {
                            $SaveVal = "Passed"   
                            $TestPass ++ 
                        }
                        "2" {
                            $SaveVal = "Failed"
                            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Alert" -Value $True
                            $TestFail ++
                        }
                        "3" {
                            $SaveVal = "Unknown"
                            $TestUnknown ++
                        }
                        Default {
                            #--[ Do nothing ]--
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "LastTestResult" -Value $SaveVal -force
                }    
                "Location" {   #--[ Location field on device must be formatted as "facility;IDF;address" separated by a semicolon ]--
                    Try{
                        $SaveVal = $Result.Value.ToString()
                    }Catch{
                        $SaveVal = ";;"
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "Facility" -Value $SaveVal.Split(";")[0] -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "IDF" -Value $SaveVal.Split(";")[1] -force
                    $Obj | Add-Member -MemberType NoteProperty -Name "Location" -Value $SaveVal.Split(";")[2] -force
                } 
                "BattRunTime" {
                    If ($Obj.HostName -Like "*chill*"){
                        $SaveVal = "N/A"
                    }Else{
                        $Result = $Result.Value.ToString()
                        $RunHours = $Result.Split(":")[0]
                        $RunMins = $Result.Split(":")[1]
                        #$RunSecs = $Result.Split(":")[2]  #--[ We don't care about seconds ]--
                        $SaveVal = $RunHours+" Hrs "+$RunMins+" Min"
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "BattRunTime" -Value $SaveVal -force
                }
                "NMC" {
                    $Flag = $False
                    $NMCArray = ($Result.Value).ToString().Split(" ")
                    ForEach ($Item in $NMCArray){
                        If ($Flag){
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCSerial" -Value $Item -force
                            $Flag = $False
                        }
                        If ($Item -like "MN:*"){
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCModelNum" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "MD:*"){
                            $String = ($Item.Split(":")[1]).Substring(0,($Item.Split(":")[1]).Length-1)
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCMfgDate" -Value $String -force
                        }
                        If ($Item -like "HR:*"){        #--[ Hardware Revision ]--                    
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCHardwareVer" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "PF:*"){        #--[ AOS Version ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSVer" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "AF1:*"){       #--[ Application Version ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAppVer" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "PN:*"){        #--[ AOS Firmware File Version ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAOSFirmware" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "AN1:*"){       #--[ Application Firmware File Version ]--
                            $Obj | Add-Member -MemberType NoteProperty -Name "NMCAppFirmware" -Value $Item.Split(":")[1] -force
                        }
                        If ($Item -like "SN:*"){
                            $Flag = $True
                        }
                    }
                }
                Default {   #--[ Use values pulled from SNMP for all others ]--
                    If (($Result.Value -eq "") -or ($Null -eq $Result.Value)){
                        $SaveVal = "existing"
                    }Else{
                        $SaveVal = $Result.Value.ToString()
                    }                  
                    If ($SaveVal -like "NoSuch*"){
                        $SaveVal = ""
                    }ElseIf ($Item -like "*date*"){    #--[ Set dates to a uniform format ]--
                        If ($SaveVal -eq ""){
                            $SaveVal = "existing"
                        }Else{
                            Try {
                                $SaveVal = Get-Date $Result.Value.ToString() -Format MM/dd/yyyy -ErrorAction SilentlyContinue
                            }Catch{ 
                                $SaveVal = $Result.Value.ToString()                   
                            }   
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name $Item[0] -Value $SaveVal -force
                }            
            }        
        }    

        #--[ Adjustments ]------------------------
        If ($Obj.HostName -like "*chill*"){
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSModelNum" -Value $Obj.NMCModelNum -force
            # $Obj | Add-Member -MemberType NoteProperty -Name "UPSSerial" -Value $Obj.NMCSerial -force  #--[ unverified ]--
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSSerial" -Value "Unknown" -force  
            # $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value ([int]((New-TimeSpan -Start ([datetime]$Obj.NMCMfgDate) -End $Today).days/365)) -force         
            $Obj | Add-Member -MemberType NoteProperty -Name "LastTestResult" -Value "N/A" -force
        }
        If ($Obj.UPSModelName -like "*Symmetra*"){
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSModelNum" -Value $Obj.UPSModelName -force
        }
        If (($Obj.UPSModelName -like "*Smart-UPS*") -or ($Obj.UPSModelName -like "*Symmetra*") -or ($Obj.UPSModelName -like "*InRow*")){   #-[ Since most don't respond with Mod # & Mfg, fake it ]--
            $Obj | Add-Member -MemberType NoteProperty -Name "Mfg" -Value "APC" -force
        }            

        #--[ Get UPS Age by MFG Date ]--
        If ($Obj.UPSMfgDate -eq ""){
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value "Unknown" -force
        }Else{
            #$MfgYear = $Obj.UPSSerial.Substring(2,3)
            #$MfgWeek = $Obj.UPSSerial.Substring(4,2)
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value ([int]((New-TimeSpan -Start ([datetime]$Obj.UPSMfgDate) -End $Today).days/365)) -force         
        }
    }Else{
        If ($Obj.Excluded){
            $Excluded ++   
            StatusMsg "Target is on exclusion list.  Bypassing..." "Cyan" $ExtOption
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Excluded" -force
        }Else{
            $Offline ++   
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Offline" -force
            If ($ExtOption.ConsoleState){Write-host "--- OFFLINE ---" -foregroundcolor "Red"}
        }
    }

    If (($ExtOption.ConsoleState) -and (!($Obj.Excluded))){
        Write-host ""
        StatusMsg "Console Mode enabled, Displaying Results..." "Magenta" $ExtOption
        $Obj
    }    

    #--[ Add data line to HTML report ]--
    $HtmlData = '<tr>'
    $HtmlData += '<td>'+$Obj.HostName+'</td>'
    $HtmlData += '<td>'+$Obj.IPAddress+'</td>'
    $HtmlData += '<td>'+$Obj.Facility+'</td>'  

    Switch ($obj.Connection){
        "OffLine"{
            $HtmlData += '<td><strong><font color=red>Offline</strong></font></td>'
        }
        "Excluded"{
            $HtmlData += '<td><strong><font color=orange>Excluded</strong></font></td>'
        }
        Default {
            $HtmlData +=  '<td><strong><font color=Green>'+$Obj.Connection+'</strong></font></td>'
        }
    }  
    
    $HtmlData += '<td>'+$Obj.MFG+'</td>'
    $HtmlData += '<td>'+$Obj.UPSModelNum+'</td>'
    $HtmlData += '<td>'+$Obj.UPSModelName+'</td>'
    $HtmlData += '<td>'+$Obj.UPSSerial+'</td>'
    $HtmlData += '<td>'+$Obj.UPSAge+'</td>'
    $HtmlData += '<td>'+$Obj.LastTestDate+'</td>'
    If ($Obj.LastTestResult -eq "Passed"){
        $HtmlData += '<td><strong><font color=Green>' 
    }ElseIf ($Obj.LastTestResult -eq "Failed"){
        $HtmlData += '<td><strong><font color=Red>'
    }Else{
        $HtmlData += '<td><strong><font color=Orange>'
    }
    $HtmlData += $Obj.LastTestResult+'</strong></font></td>'
    $HtmlData += '<td>'+$Obj.BattRunTime+'</td>'
    $HtmlData += '</tr>'
    $HtmlReport += $HtmlData   
    $Obj = ""   
}

#--[ HTML Email Report ]--

If ($TestFail -gt 0){
    $Title = 'darkred><strong>!! Battery Backup Status Report !!'
}ElseIf (($Offline -gt 0) -or ($TestUnknown -gt 0)){
    $Title = 'orange><strong>!! Battery Backup Status Report !!'
}Else{
    $Title = 'darkcyan><strong>Battery Backup Status Report'
}

$Columns = 12  #--[ Total columns in report ]--
$HtmlHeader += '
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
</head>
'
$HtmlBody +='
<body>
    <div class="content">
    <table border-collapse="collapse" border="3" cellspacing="0" cellpadding="5" width="100%" bgcolor="#E6E6E6" bordercolor="black">
        <tr>
            <td colspan='+$Columns+'><center><H2><font color='+$Title+'</strong></H2></center></td>
        </tr>
        <tr>
            <td colspan='+$Columns+'>
                <center>
                    <table border-collapse="collapse" border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#E6E6E6" bordercolor="black">
                        <td border="0"><strong><center>Total Devices = '+$Count+'</center></td>'
                        If ($Offline -gt 0){
                            $HtmlBody +='<td><font color="red"><strong><center>Offline = '+$Offline+'</center></td>'
                        }Else{
                            $HtmlBody +='<td><font color="green"><strong><center>Offline = '+$Offline+'</center></td>'
                        }
                        $HtmlBody +='<td><font color="green"><strong><center>Self-Test Passes = '+$TestPass+'</center></td>'
                        If ($TestFail -gt 0){
                            $HtmlBody +='<td><font color="red"><strong><center>Self-Test Failures = '+$TestFail+'</center></td>'
                        }Else{
                            $HtmlBody +='<td><font color="green"><strong><center>Self-Test Failures = '+$TestFail+'</center></td>'
                        }
                        If ($TestUnknown -gt 0){                     
                            $HtmlBody +='<td><font color="orange"><strong><center>Unknown Status = '+$TestUnknown+'</center></td>'
                        }Else{
                            $HtmlBody +='<td><font color="green"><strong><center>Unknown Status = '+$TestUnknown+'</center></td>'
                        }
                        $HtmlBody +='
                    </table>
                </center>
            </td>
        </tr>
        <tr>
            <td><strong><center>Host Name</center></td>
            <td><strong><center>IP Address</center></td>
            <td><strong><center>Location</center></td>
            <td><strong><center>Status</center></td>
            <td><strong><center>Mfg</center></td>
            <td><strong><center>Model #</center></td>
            <td><strong><center>Model Name</center></td>
            <td><strong><center>Serial</center></td>
            <td><strong><center>Age (Years)</center></td>
            <td><strong><center>Last Test</center></td>
            <td><strong><center>Result</center></td>
            <td><strong><center>Runtime</center></td>
        </tr>
'

#--[ Construct final full report ]--
$DateTime = Get-Date -Format MM-dd-yyyy_hh:mm:ss 
$HtmlReport = $HtmlHeader+$HtmlBody+$HtmlReport
$HtmlReport += '<tr><td colspan='+$Columns+'><center><font color=darkcyan><strong>Audit completed at: '+$DateTime+'</strong></center></td></tr>'   
$HtmlReport += '</table></div></body></html>'

#--[ Only keep the last 10 of the log files ]-- 
Try{
    If (!(Test-Path -PathType container ($SourcePath+"\Reports"))){
          New-Item -ItemType Directory -Path ($PSScriptRoot+"\Reports") -Force
    }
    Get-ChildItem -Path ($SourcePath+"\Reports") | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*html*")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item | Out-Null
    $DateTime = Get-Date -Format MM-dd-yyyy_hh.mm.ss 
    $Report = ($SourcePath+"\Reports\UPS-Status_"+$DateTime+".html")
    StatusMsg $Report "Yellow" $ExtOption
    Add-Content -Path $Report -Value $HtmlReport#>
}Catch{
    StatusMsg "Folder access issue.  Logfile cannot be saved." "Red" $ExtOption
    $HtmlReport += "NOTE:  There is an issue with folder access.  Log file has NOT been saved."
}

#--[ Set the alternate email recipient if running out of an IDE console for testing ]-- 
If ($Null -ne $ExtOption.EmailAltRecipient){
    #--[ Use the setting from XML ]--
}Else{
    If ($Env:Username.SubString(0,1) -eq "a"){       #--[ Filter out admin accounts ]--
        $ThisUser = ($Env:Username.SubString(1))+"@"+$Env:USERDNSDOMAIN 
        $ExtOption | Add-Member -MemberType NoteProperty -Name "EmailAltRecipient" -Value $ThisUser -force
    }Else{
        $ThisUser = $Env:USERNAME+"@"+$Env:USERDNSDOMAIN 
        $ExtOption | Add-Member -MemberType NoteProperty -Name "EmailAltRecipient" -Value $ThisUser -force
    }
}

SendEmail $HtmlReport $ExtOption 

#--[ Use this to load the report in the default browser ]--
# iex $Report

If ($ExtOption.ConsoleState){Write-host "`n--- Completed ---" -foregroundcolor red}

#}catch{
 #   add-content -Path $PSScriptRoot\new-error.log -Value $_.Exception.Message
#}
<#--[ XML File Example -- File should be named same as the script ]--
<!-- Settings & configuration file -->
<Settings>
    <General>
        <SourcePath>c:\Scripts\UPS-Inventory</SourcePath>
        <ExcelSourceFile>UPS-Inventory.xlsx</ExcelSourceFile>
		<DNS>10.1.1.1</DNS>
        <IPListFile>IP-List.txt</IPListFile>
        <SmtpServer>mail.company.org</SmtpServer>
        <EmailRecipient>it@company.org</EmailRecipient>
        <EmailAltRecipient></EmailAltRecipient>  <!-- If blank the user ID running the script is used -->
        <EmailSender>ups@company.org</EmailSender>
   		<HNPattern>UPS</HNPattern>
   		<DayOfWeek>Sunday</DayOfWeek>
        <AlwaysAlt>True</AlwaysAlt>  <!-- When set to TRUE the alt email recipient gets an email at every run -->
    </General>
    <Exclusions>
		<Exclude>10.10.10.21,10.10.120.22,10.10.12.23</Exclude>
	</Exclusions>
    <Credentials>
    	<PasswordFile>passfile.txt</PasswordFile>
	    <KeyFile>c:\keyfile.txt</KeyFile>
	    <SNMPv3User>snmpv3user</SNMPv3User>
        <SNMPv3AltUser>snmpv3altusername</SNMPv3AltUser>
		<SNMPv3Secret>bahbahblacksheep</SNMPv3Secret>  
        <SNMPv3Secret>mysnmp3pass</SNMPv3Secret>  
    </Credentials>
</Settings>    
  
#>

<#--[ APC OID Reference ]--
    "BattFreqOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.2.0"
    "BattVOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.1.0"
    "BattVIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.1.0"
    "BattFreqIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.4.0"
    "BattActualV" = ".1.3.6.1.4.1.318.1.1.1.2.2.8.0"
    "BattCurrentAmps" = ".1.3.6.1.4.1.318.1.1.1.2.2.9.0"
    "BattLastRepl" = ".1.3.6.1.4.1.318.1.1.1.2.1.3.0"
    "BattCapLeft" = ".1.3.6.1.4.1.318.1.1.1.2.2.1.0"
    "BattRunTime" = ".1.3.6.1.4.1.318.1.1.1.2.2.3.0"
    "BattReplace" = ".1.3.6.1.4.1.318.1.1.1.2.2.4.0"
    "BattTemp" = ".1.3.6.1.4.1.318.1.1.1.2.2.2.0"
    "ACVIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.1.0"
    "ACFreqIn" = ".1.3.6.1.4.1.318.1.1.1.3.2.4.0"
    "LastXfer" = ".1.3.6.1.4.1.318.1.1.1.3.2.5.0"
    "UPSVOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.1.0"
    "UPSFreqOut" = ".1.3.6.1.4.1.318.1.1.1.4.2.2.0"
    "UPSOutLoad" = ".1.3.6.1.4.1.318.1.1.1.4.2.3.0"
    "UPSOutAmps" = ".1.3.6.1.4.1.318.1.1.1.4.2.4.0"
    "LastTestResult" = ".1.3.6.1.4.1.318.1.1.1.7.2.3.0"
    "LastTestDate" = ".1.3.6.1.4.1.318.1.1.1.7.2.4.0"
    "BIOSSerNum" = ".1.3.6.1.4.1.318.1.1.1.1.2.3.0"
    "UPSModel" = ".1.3.6.1.4.1.318.1.1.1.1.1.1.0"
    "FirmwareVer" = ".1.3.6.1.4.1.318.1.1.1.1.2.1.0"
    "MfgDate" = ".1.3.6.1.4.1.318.1.1.1.1.2.2.0"
    "Location" = ".1.3.6.1.2.1.1.6.0"
    "Contact" = ".1.3.6.1.2.1.1.4.0"       
#>

