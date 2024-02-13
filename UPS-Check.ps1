Param(
    [switch]$Console = $false,                                                  #--[ Set to true to enable local console result display. Defaults to false ]--
    [switch]$Debug = $False,                                                    #--[ Generates extra console output for debugging.  Defaults to false ]--
    [switch]$EnableExcel = $True                                                #--[ Defaults to use Excel. ]--  
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
                   : Currently designed to poll APC UPS devices.  UPS NMC must have SNMP v3 active.
                   : Script checks for active SNMPv1, FTP, and SNMPv3.
                   : Will generate a new spreadsheet if none exists by using a text file located in the same folder
                   : as the script, one IP per line.  Default operation is to check for text file first, then if not
                   : found check for an existing spreadsheet also in the same folder.  If an existing spreadsheet
                   : is located the target list is compliled from column A.  It will also copy a master spreadsheet
                   : to a working copy that gets processed.  Up to 10 backup copies are retained prior to writing
                   : changes to the working copy.
                   : 
          Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is running or it can crash.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   :   not limited to the following:
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.0 - 08-16-22 - Original 
    Change History : v2.0 - 09-00-22 - Numerous operational & bug fixes prior to v3.
                   : v3.1 - 09-15-22 - Cleaned up final version for posting.
                   : v4.0 - 04-12-23 - Too many changes to list
                   : v4.1 - 07-03-23 - Added age and LDOS dates. 
                   : v5.0 - 01-17-24 - Fixed DNS lookup.  Fixed last test result.  Fixed color coding of hostname for
                   :                   numerous events.  Added hostname cell comments to describe color coding.
                   : v6.0 - 02-12-24 - Retooled Html email report.  Added self test failed counts.  Added saved reports.
                   : v6.1 - 02-13-24 - Added missing external config entries.
                   :                   
==============================================================================#>

#Requires -version 5
##equires -Modules @{ ModuleName="SNMPv3"; ModuleVersion="1.1.1" }
Clear-Host 

#--[ Variables ]---------------------------------------------------------------
$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$Today = Get-Date -Format MM-dd-yyyy 
$Script:v3UserTest = $False

#--[ Runtime tweaks for testing ]--
$EnableExcel = $True
$Console = $True
$Debug = $false #true
$CloseOpen = $true

#------------------------------------------------------------
#
$erroractionpreference = "stop"
try{
    if (!(Get-Module -Name SNMPv3)) {
        Get-Module -ListAvailable SNMPv3 | Import-Module | Out-Null
    }
    Install-Module -Name SNMPv3
}Catch{

}

#==[ Functions ]===============================================================
Function SendEmail ($MessageBody,$ExtOption) {    
    $ErrorActionPreference = "Stop"
    $Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,25) 
    Try{  
        If ($Env:Username.SubString(0,1) -eq "a"){
            $ThisUser = ($Env:Username.SubString(1))+"@"+$Env:USERDNSDOMAIN 
        }Else{
            $ThisUser = $Env:USERNAME+"@"+$Env:USERDNSDOMAIN 
        }
    }Catch{
        $ThisUser = "UPS_Status"
    }
    $Email = New-Object System.Net.Mail.MailMessage  
    $Email.IsBodyHTML = $true
    $Email.From = $ThisUser
    If ($ExtOption.ConsoleState){  #--[ If running out of an IDE console, send to the user only for testing ]--
        $ThisUser = $Env:USERNAME+"@"+$Env:USERDNSDOMAIN 
        $Email.To.Add($ThisUser) 
    }Else{
        If ($ExtOption.Recipient -eq ""){
            $Email.To.Add($ThisUser) 
        }Else{
            $Email.To.Add($ExtOption.EmailRecipient) 
        }
    }
    $Email.Subject = "UPS Status Report"
    $Email.Body = $MessageBody
    Try {
        $Smtp.Send($Email)
        If ($ExtOption.ConsoleState){Write-Host `n"--- Email Sent ---" -ForegroundColor red }
    }Catch{
        $_.Error.Message
        $_.Exception.Message
    }
}
Function StatusMsg ($Msg, $Color, $ExtOption){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    $Msg = ""
}

Function LoadConfig ($ExtOption,$ConfigFile){  #--[ Read and load configuration file ]-------------------------------------
    StatusMsg "Loading external config file..." "Magenta" $ExtOption
    if (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
        [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--    
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SourcePath" -Value $Config.Settings.General.SourcePath
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelSourceFile" -Value $Config.Settings.General.ExcelSourceFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DNS" -Value $Config.Settings.General.DNS
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SMNPv3User" -Value $Config.Settings.Credentials.SMNPv3User
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SMNPv3AltUser" -Value $Config.Settings.Credentials.SMNPv3AltUser
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SNMPv3Secret" -Value $Config.Settings.Credentials.SMNPv3Secret
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PasswordFile" -Value $Config.Settings.Credentials.PasswordFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "KeyFile" -Value $Config.Settings.Credentials.KeyFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "IPlistFile" -Value $Config.Settings.General.IPListFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.General.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $Config.Settings.General.EmailRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "HNPattern" -Value $Config.Settings.General.HNPattern
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $ExtOption
        break;break;break
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

Function SMNPv3Walk ($Target,$OID,$Debug){
    $WalkRequest = @{
        UserName   = $Script:SMNPv3User
        Target     = $Target
        OID        = $OID
        AuthType   = 'MD5'
        AuthSecret = $Script:SNMPv3Secret
        PrivType   = 'DES'
        PrivSecret = $Script:SNMPv3Secret
        #Context    = ''
    }
    $Result = Invoke-SNMPv3Walk @WalkRequest | Format-Table -AutoSize
    If ($Debug){write-host "SNMpv3 Debug :" $Result }
    Return $Result
}

Function GetSNMPv1 ($Target,$OID,$Debug) {
    $SNMP = New-Object -ComObject olePrn.OleSNMP
    $erroractionpreference = "Stop"
    Try{
        $snmp.open($Target,$Script:SMNPv3User,2,1000)
        $Result = $snmp.get($OID)
    }Catch{
        Return $_.Exception.Message        
    }
    If ($Debug){ write-host "SNMpv1 Debug :" $Result }
    Return $Result
    #$snmp.gettree('.1.3.6.1.2.1.1.1')
}

Function GetSMNPv3 ($Target,$OID,$Debug,$Test){
    If ($Test){  #--[ If 1st user tests positive on 1st use, use it by setting the global variable below ]--
        $GetRequest1 = @{
            UserName   = $Script:SMNPv3User
            Target     = $Target
            OID        = $OID.Split(",")[1]
            AuthType   = 'MD5'
            AuthSecret = $Script:SNMPv3Secret
            PrivType   = 'DES'
            PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest1 -ErrorAction:Stop
            If ($Result -like "*Exception*"){
                $Script:v3UserTest = $False  
            }Else{
                $Script:v3UserTest = $True  #--[ Global v3 user variable ]--
            }
        }Catch{
            If ($Debug){
                write-host $_.Exception.Message -ForegroundColor Cyan
                write-host " -- SNMPv3 User 1 failed..." -ForegroundColor red
            }
        }
    }Else{  #--[ User 1 has failed so use user 2 instead ]--
        $GetRequest2 = @{
            UserName   = $Script:SMNPv3AltUser
            Target     = $Target
            OID        = $OID.Split(",")[1]
            #AuthType   = 'MD5'
            #AuthSecret = $Script:SNMPv3Secret
            #PrivType   = 'DES'
            #PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest2 -ErrorAction:Stop
        }Catch{
            If ($Result -like "*Exception*"){
                write-host " -- SNMPv3 User 2 failed... No SNMPv3 access..." -ForegroundColor red
                write-host $_.Exception.Message -ForegroundColor Blue
            }
        }
    }
    If ($Debug){
        Write-Host "  -- SNMPv3 Debug: " -ForegroundColor Yellow -NoNewline
        If ($Test){
            Write-Host "SNMP User 2  " -ForegroundColor Green -NoNewline
        }Else{
            Write-Host "SNMP user 1  " -ForegroundColor Green -NoNewline
        }
        Write-Host $OID.Split(",")[0]"  " -ForegroundColor Cyan -NoNewline
        Write-Host $Result    
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

StatusMsg "Processing UPS Devices" "Yellow" $ExtOption
$erroractionpreference = "stop"

#--[ Close copies of Excel that PowerShell has open ]--
If ($CloseOpen){
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
$Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
StatusMsg "Creating new Excel COM object..." "Magenta" $ExtOption

If (Test-Path -Path $ExtOption.SourcePath -PathType leaf){
    $SourcePath = $ExtOption.SourcePath
}Else{
    $SourcePath = $PSScriptRoot
}

#--[ If this file exists the IP list will be pulled from it ]--
If (Test-Path -Path ($SourcePath+"\"+$ExtOption.IPListFile) -PathType Leaf){   
    $IPList = Get-Content ($SourcePath+"\"+$ExtOption.IPListFile)         
    StatusMsg "IP text list was found, loading IP list from it... " "green"  $ExtOption
}else{ 
    StatusMsg "IP text list not found, Attempting to process spreadsheet... " "cyan" $ExtOption
    If (Test-Path -Path ($SourcePath+"\"+$ExtOption.ExcelSourceFile) -PathType Leaf){
        $Excel.Visible = $false
        $WorkBook = $Excel.Workbooks.Open(($SourcePath+"\"+$ExtOption.ExcelSourceFile))
        $WorkSheet = $Workbook.WorkSheets.Item("UPS")
        $WorkSheet.activate()       
    }Else{
        StatusMsg "Existing spreadsheet not found, Source copy failed, Nothing to process.  Exiting... " "red" $ExtOption
        Break;break
    }
    $WorkSheet = $Workbook.WorkSheets.Item("UPS")
    $WorkSheet.activate()    
    $Row = 4   
    $IPList = @() 
    Do {
        $IPList += ,@($Row,$WorkSheet.Cells.Item($row,1).Text)
        $Row++
    } Until (
        $WorkSheet.Cells.Item($row,1).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
    )
}

$Excel.DisplayAlerts = $false
$WorkBook.Close($true)
$Excel.Quit()

ForEach ($Target in $IPList){  #--[ Are we pulling from Excel or a text file?  Jagged has row numbers from Excel ]--
    if ($Target.length -eq 2){
        $Jagged = $True
    }Else{
        $Jagged = $False
    }
}

#==[ Process items collected in $IPList, from spreadsheet, or text file as appropriate ]===============================
$Row = 4   
$TestPass = 0
$TestFail = 0
$TestUnknown = 0
$Offline = 0
$HtmlHeader = @() 
$HtmlReport = @() 
$HtmlBody = @()
$Count = $IPList.Count

ForEach ($Target in $IPList){
    If ($Jagged){
        $Row = $Target[0]
        $Target = $Target[1]
    }
    $Current = $Row-3
    If ($ExtOption.ConsoleState){Write-Host "`nCurrent Target  :"$Target"  ("$Current" of "$Count")" -ForegroundColor Yellow }
   
    $Obj = New-Object -TypeName psobject   #--[ Collection for Results ]--
    Try{
        $HostLookup = (nslookup $Target $Script:DNS 2>&1) 
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value ($HostLookup[3].split(":")[1].TrimStart()).Split(".")[0] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $True
    }Catch{
        $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value "Not Found" -force
        $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False
    }
    
    $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $Target -force

    If (Test-Connection -ComputerName $Target -count 1 -BufferSize 16 -Quiet){  #--[ Ping Test ]--
        $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Online" -force
        If ($ExtOption.ConsoleState){StatusMsg "Polling SNMP..." "Magenta" $ExtOption}
        If ((!($Debug)) -and ($ExtOption.ConsoleState)){Write-Host "  Working." -NoNewline}

        #--[ Test for SNMPv3.  Make sure to include leading comma  ]---------
        $Test = GetSMNPv3 $Target ",1.3.6.1.2.1.1.8.0" $Debug $Script:v3UserTest
        if ($Test -like "*TimeTicks*"){
            $PortTest = "True"
        }Else{
            $PortTest = "False"
        }
        $Obj | Add-Member -MemberType NoteProperty -Name "SNMPv3" -Value $PortTest -force
        If ((!($Debug)) -and ($ExtOption.Console)){Write-Host "." -NoNewline}

        #--[ Test for SNMPv1 ]------------------------------------------------
        $Test = GetSNMPv1 $Target "1.3.6.1.2.1.1.8.0" $Debug
        if ($Test -like "*TimeTicks*"){
            $PortTest = "True"
        }Else{
            $PortTest = "False"
        }
        $Obj | Add-Member -MemberType NoteProperty -Name "SNMPv1" -Value $PortTest -force
    }

    #--[ Only process OIDs if online PLUS SNMPv3 is good ]--------------------------
    If (($Obj.Connection -eq "Online") -and ($Obj.SNMPv3 -ne "False")){  
        ForEach ($Item in $OIDArray){            
            $Result = GetSMNPv3 $Target $Item $Debug $Script:v3UserTest
            If ($Debug){
                Write-Host ' '$Item[0]'='$Result -ForegroundColor yellow
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
                        }
                        "2" {
                            $SaveVal = "Failed"
                            $Script:Alert = $True 
                        }
                        "3" {
                            $SaveVal = "Unknown"
                        }
                        Default {
                            $SaveVal = "Unknown"
                        }
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name "LastTestResult" -Value $SaveVal -force
                }    
                "Location" {   #--[ Location field on device must be formatted as "facility;IDF;address" separated by a semicolon ]--
                    $SaveVal = $Result.Value.ToString()
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
                        #$RunSecs = $Result.Split(":")[2]
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
                }   #>         
            }        
        }    
        #--[ Adjustments ]------------------------
        If ($Obj.HostName -like "*chill*"){
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSModelNum" -Value $Obj.NMCModelNum -force
            # $Obj | Add-Member -MemberType NoteProperty -Name "UPSSerial" -Value $Obj.NMCSerial -force  #--[ unverified ]--
            $Obj | Add-Member -MemberType NoteProperty -Name "UPSSerial" -Value "Unknown" -force  
            # $Obj | Add-Member -MemberType NoteProperty -Name "UPSAge" -Value ([int]((New-TimeSpan -Start ([datetime]$Obj.NMCMfgDate) -End $Today).days/365)) -force         
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
        $Offline ++   
        $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Offline" -force
        If ($ExtOption.ConsoleState){Write-host "--- OFFLINE ---" -foregroundcolor "Red"}
    }

    If ($ExtOption.ConsoleState){
        Write-host " "
        $Obj
    }    

    #--[ Add data line to HTML report ]--
    $HtmlData = '<tr>'
    $HtmlData += '<td>'+$Obj.HostName+'</td>'
    $HtmlData += '<td>'+$Obj.IPAddress+'</td>'
    $HtmlData += '<td>'+$Obj.Facility+'</td>'  
    If ($obj.Connection -eq "OffLine"){
        $HtmlData += '<td><strong><font color=red>' #+$Obj.Connection+'</strong></font></td>'
    }Else{
        $HtmlData += '<td><strong><font color=green>'
    }
    $HtmlData += $Obj.Connection+'</strong></font></td>'
    
    $HtmlData += '<td>'+$Obj.MFG+'</td>'
    $HtmlData += '<td>'+$Obj.UPSModelNum+'</td>'
    $HtmlData += '<td>'+$Obj.UPSModelName+'</td>'
    $HtmlData += '<td>'+$Obj.UPSSerial+'</td>'
    $HtmlData += '<td>'+$Obj.UPSAge+'</td>'
    $HtmlData += '<td>'+$Obj.LastTestDate+'</td>'
    If ($Obj.LastTestResult -eq "Passed"){
        $HtmlData += '<td><strong><font color=Green>' #+$Obj.LastTestResult+'</strong></font></td>'
        $TestPass ++
    }ElseIf ($Obj.LastTestResult -eq "Failed"){
        $HtmlData += '<td><strong><font color=Red>'
        $TestFail ++
    }Else{
        $HtmlData += '<td><strong><font color=Orange>'
        $TestUnknown ++
    }
    $HtmlData += $Obj.LastTestResult+'</strong></font></td>'
    $HtmlData += '<td>'+$Obj.BattRunTime+'</td>'
    $HtmlData += '</tr>'
    $HtmlReport += $HtmlData      
}

#--[ HTML Email Report ]--
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
            <td colspan='+$Columns+'><center><H2><font color=darkcyan><strong>Battery Backup Status Report</strong></H2></center></td>
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
            $HtmlBody +='<td><font color="green"><strong><center>Self Test Passes = '+$TestPass+'</center></td>'
If ($TestFail -gt 0){
            $HtmlBody +='<td><font color="red"><strong><center>Self Test Failures = '+$TestFail+'</center></td>'
}Else{
            $HtmlBody +='<td><font color="green"><strong><center>Self Test Failures = '+$TestFail+'</center></td>'
}
If ($TestUnknown -gt 0){                     
            $HtmlBody +='<td><font color="orange"><strong><center>Unknow Status = '+$TestUnknown+'</center></td>'
}Else{
            $HtmlBody +='<td><font color="green"><strong><center>Unknow Status = '+$TestUnknown+'</center></td>'
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
If (!(Test-Path -PathType container ($SourcePath+"\Reports"))){
      New-Item -ItemType Directory -Path ($PSScriptRoot+"\Reports") -Force
}
Get-ChildItem -Path ($SourcePath+"\Reports") | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*html*")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item
$DateTime = Get-Date -Format MM-dd-yyyy_hh.mm.ss 
$Report = ($SourcePath+"\Reports\UPS-Status_"+$DateTime+".html")
Add-Content -Path $Report -Value $HtmlReport

#--[ Use to load the report in the default browser ]--
# iex "$PSScriptRoot\temp.html"

SendEmail $HtmlReport $ExtOption 
If ($ExtOption.ConsoleState){Write-host "`n--- Completed ---" -foregroundcolor red}


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
    </General>
    <Credentials>
	<PasswordFile>passfile.txt</PasswordFile>
	<KeyFile>c:\keyfile.txt</KeyFile>
	<SMNPv3User>snmpv3user</SMNPv3User>
        <SMNPv3AltUser>snmpv3altusername</SMNPv3AltUser>
	<!--	<SNMPv3Secret>bahbahblacksheep</SNMPv3Secret>  -->
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

