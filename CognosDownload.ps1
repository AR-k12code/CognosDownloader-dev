Param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Give the name of the report you want to download.")][string]$report,
    [parameter(Position=1,Mandatory=$false,HelpMessage="Give a specific folder to download the report into.")][string]$savepath="C:\ImportFiles\", #--- VARIABLE --- change to a path you want to save files
    [parameter(Position=2,Mandatory=$false,HelpMessage="Extension to save onto the report name.")][ValidateSet("csv","xlsx")][string]$extension="csv", #--- VARIABLE --- file extension to save data in csv or xlsx
    [parameter(Mandatory=$false,HelpMessage="eSchool SSO username to use.")][string]$username="0000name", #--- VARIABLE --- SSO username
    [parameter(Mandatory=$false,HelpMessage="File for ADE SSO Password")][string]$passwordfile="C:\Scripts\apscnpw.txt", #--- VARIABLE --- change to a file path for SSO password
    [parameter(Mandatory=$false,HelpMessage="eSchool DSN location.")][string]$espdsn="schoolsms", #--- VARIABLE --- eSchool DSN for your district
    [parameter(Mandatory=$false,HelpMessage="eFinance username to use.")][string]$efpuser="yourefinanceusername", #--- VARIABLE --- eFinance username
    [parameter(Mandatory=$false,HelpMessage="eFinance DSN location.")][string]$efpdsn="schoolfms", #--- VARIABLE --- eFinance DSN for your district
    [parameter(Mandatory=$false,HelpMessage="Cognos Folder Structure.")][string]$cognosfolder="My Folders", #--- VARIABLE --- Cognos Folder "Folder 1/Sub Folder 2/Sub Folder 3" NO TRAILING SLASH
    [parameter(Mandatory=$false,HelpMessage="Report Parameters")][string]$reportparams="", #--- VARIABLE --- Example:"p_year=2017&p_school=Middle School" If a report requires parameters you can specifiy them here.
    [parameter(Mandatory=$false,HelpMessage="Report Wait Timeout")][int]$reportwait=5, #--- VARIABLE --- If the report is not ready immediately wait X seconds and try again. Will try 6 times only!
    [parameter(Mandatory=$false,HelpMessage="Use switch for Report Studio created report. Otherwise it will be a Query Studio report")][switch]$ReportStudio,
    [parameter(Mandatory=$false,HelpMessage="Get the report from eFinance.")][switch]$eFinance,
    [parameter(Mandatory=$false,HelpMessage="Run a live version instead of just getting a saved version.")][switch]$RunReport,
    [parameter(Mandatory=$false,HelpMessage="Send an email on failure.")][switch]$SendMail,
    [parameter(Mandatory=$false,HelpMessage="SMTP Auth Required.")][switch]$smtpauth,
    [parameter(Mandatory=$false,HelpMessage="SMTP Server")][string]$smtpserver="smtp-relay.gmail.com", #--- VARIABLE --- change for your email server
    [parameter(Mandatory=$false,HelpMessage="SMTP Server Port")][int]$smtpport="587", #--- VARIABLE --- change for your email server
    [parameter(Mandatory=$false,HelpMessage="SMTP eMail From")][string]$mailfrom="noreply@yourdomain.com", #--- VARIABLE --- change for your email from address
    [parameter(Mandatory=$false,HelpMessage="File for SMTP eMail Password")][string]$smtppasswordfile="C:\Scripts\emailpw.txt", #--- VARIABLE --- change to a file path for email server password
    [parameter(Mandatory=$false,HelpMessage="Send eMail to")][string]$mailto="technology@yourdomain.com", #--- VARIABLE --- change for your email to address
    [parameter(Mandatory=$false,HelpMessage="Minimum line count required for CSVs")][int]$requiredlinecount=3, #This should be the ABSOLUTE minimum you expect to see. Think schools.csv for smaller districts.
    [parameter(Mandatory=$false)][switch]$showURL #print the url used to download report.
)

Add-Type -AssemblyName System.Web

# The above parameters can be called directly from powershell switches
# In Cognos, do the following:
# 1. Setup a report with specific name (best without spaces like MyReportName) to run scheduled to save with which format you want then schedule this script to download it.
# 2. You will need to determine the DSN (database name) for your district
#    The database name is typically YOURSCHOOLNAMEsms (eSchool Plus) or YOURSCHOOLNAMEfms (eFinance)
#    Click the Login Avatar in the top right and click Sign In. Select esp (eSchool Plus) or efp (eFinance) then find your school database in the list.
# On computer to download data:
# 1. After adjusting relevant variables for your district and user account
# 2. Create folder to store password data for script (default of C:\scripts)
# 3. Run from command line or batch script: powershell.exe -executionpolicy bypass -file C:\Scripts\CognosDownload.ps1 MyReportName
#    Scripts can also use command switches from powershell: CognosDownload.ps1 -report MyReportName -savepath C:\Scripts -extension csv -username 0000username -espdsn schoolsms
# In case report was not built with Query Studio, use -ReportStudio
# For eFinance:
# Use "-username 0000name -efpuser yourefinanceusername -efpdsn schoolfms -eFinance" to run report from eFinance (SSO username is required along with eFinance username)

# When the password expires, just delete the specific file and run the script to re-create

$userdomain = "APSCN"
#******************* end of variables to change ********************

# Revisions:
# 2014-07-23: Brian Johnson: Updated URL string to include dsn parameters necessary for eSchool and re-enabled CredentialCache setting to login
# 2016-04-06: Added new username parameter efpuser for eFinance to work
# 2017-01-16: Brian Johnson: Updated URL from cognosisapi.dll to cognos.cgi. Also included previous changes that were not uploaded from before.
# 2017-02-07: Added CSV verify and revert
# 2017-02-27: Added variable for reporttype
# 2017-07-12: VBSDbjohnson: Merged past changes with CWeber42 version
# 2017-07-13: VBSDbjohnson: Changed to use Powershell parameters instead of args. Script should also be able to run without modifying file
# 2018-04-26: (reverted) scottorgan: Nested folder support Usage examples: CognosDownload.ps1 Clever\Entollments ; CognosDownload.ps1 "Other Reports\MAP Roster"
# 2018-04-26: (reverted) BPSDJreed: Email notification for expired password
# 2018-04-24: Craig Millsap: Added recursive nested folders, email notifications, waiting for report to generate.
# 2020-11-19: Craig Millsap: Major overhaul for Cognos11 upgrade. Working eSchool Plus. eFinance is not working yet.

#send mail on failure.
$mailsubject = "[CognosDownloader]"
function Send-Email([string]$failurereason,[string]$errormessage) {
    if ($SendMail) {
        $msg = New-Object Net.Mail.MailMessage
        $smtp = New-Object Net.Mail.SmtpClient($smtpserver, $smtpport)
        #port 25 is likely non-ssl (for internal restricted relays), maybe change to switch option?
        if ($smtpport -eq 25) {$smtp.EnableSSL = $False} else { $smtp.EnableSSL = $True }
        #If authentication is required.
        if ($smtpauth) { $smtp.Credentials = New-Object System.Net.NetworkCredential($mailfrom,$mailfrompassword) }
        $msg.From = $mailfrom
        $msg.To.Add($mailto)
        #Include date so emails don't group in a thread.
        $msg.subject =  $mailsubject + $failurereason + "[$(Get-Date -format MM/dd/y)]" + '[' + $report + ']'
        $msg.Body = "The report " + $report  + " failed to download properly.`r`n"
        if ($errormessage) {
            $msg.Body += "$errormessage`r`n"
        }
        $msg.Body += $url
        
        try {
            $smtp.send($msg)
        } catch {
            Write-Host("Failed to send email: $_") -ForeGroundColor Red
            exit 30
        }
    }
}

function Reset-DownloadedFile([string]$fullfilepath) {
    $PrevOldFileExists = Test-Path ($fullfilepath + ".old")
    if ($PrevOldFileExists -eq $True) {
        Write-Host("Deleting old $report...") -ForeGroundColor Yellow
        Remove-Item -Path $fullfilepath -Force
        Rename-Item -Path ($fullfilepath + ".old") -newname ($fullfilepath)
    }
    Write-Host("Failed CSV verify. Reversing old $report...") -ForeGroundColor Red
}

# Cognos ui action to perform 'run' or 'view'
# run not fully implemented
If ($RunReport) {$uiAction = "run"} Else {$uiAction = "view"}

# server location for Cognos
$baseURL = "https://dev.adecognos.arkansas.gov"
$cWebDir = "ibmcognos"

If ($eFinance) {
    $camName = "efp"    #efp for eFinance
    $camuser = $efpuser
    $dsnparam = "spi_db_name"
    $dsnname = $efpdsn
} else {
    $camName = "esp"    #esp for eSchool
    $camuser = $username
    $dsnparam = "dsn"
    $dsnname = $espdsn
}

#report for Report Studio, query for Query Studio
if ($ReportStudio) { 
    $reporttype = "report"
} else {
    $reporttype = "query"
}

#Script to create a password file for Cognos download Directory
#This script MUST BE RAN LOCALLY to work properly! Run it on the same machine doing the cognos downloads, this does not work remotely!

if ((Test-Path ($passwordfile))) {
    $password = Get-Content $passwordfile | ConvertTo-SecureString
} else {
    Write-Host("Password file does not exist! [$passwordfile]. Please enter a password to be saved on this computer for scripts") -ForeGroundColor Yellow
    Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $passwordfile
    $password = Get-Content $passwordfile | ConvertTo-SecureString
}

$creds = New-Object System.Management.Automation.PSCredential $username,$password

If ($smtpauth) {
    if ((Test-Path ($smtppasswordfile))) {
        $smtppassword = Get-Content $smtppasswordfile | ConvertTo-SecureString
    } else {
        Write-Host("SMTP Password file does not exist! [$smtppasswordfile]. Please enter a password to be saved on this computer for emails") -ForeGroundColor Yellow
        Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $smtppasswordfile
        $mailfrompassword = Get-Content $smtppasswordfile | ConvertTo-SecureString
    }
}

switch ($extension) {
    "csv" { $fileformat = "CSV" }
    "xlsx" { $fileformat = "spreadsheetML" }
    DEFAULT { $fileformat = "CSV" }
}

$fullfilepath = "$savepath\$report.$extension"

If (!(Test-Path ($savepath))) {
    Write-Host("Specified save folder does not exist! [$fullfilepath]") -ForeGroundColor Yellow
    Send-Email("[Failure][Save Path Missing]")
    exit 1 #specified save folder does not exist
}

#get current datetime for if-modified-since header for file
$filetimestamp = Get-Date

#Include folders. In the future create an array from split at /
if ($cognosfolder -eq "My Folders") {
    $cognosfolder = "folder[@name='My Folders']/"
    $cognosfolder = $([System.Web.HttpUtility]::UrlEncode($cognosfolder)).replace('+','%20')
} else {
    $folders = $cognosfolder.split('/')
    $cognosfolder = ''
    for ($counter=0; $counter -lt $folders.Length; $counter++) {
        $cognosfolder += "folder[@name='" + [string]$folders[$counter] + "']/"
    }
    $cognosfolder = "folder[@name='My Folders']/" + $cognosfolder
    $cognosfolder = $([System.Web.HttpUtility]::UrlEncode($cognosfolder)).replace('+','%20')
}

#Write-Host $cognosfolder
#Write-Host $([System.Web.HttpUtility]::UrlDecode($cognosfolder))

if ($reportparams.Length -gt 0) {
    #$reportparams = $([System.Web.HttpUtility]::UrlEncode('&' + $reportparams)).replace('+','%20').replace('%26','&')
    $reportparams = '&' + $reportparams
}

$camid = "CAMID(%22$($camName)%3aa%3a$($camuser)%22)%2f$($cognosfolder)$($reporttype)%5b%40name%3d%27$($report)%27%5d"

if ($uiAction -match "run") { #run the report live for the data
    $url = "$($baseURL)/$($cWebDir)/bi/v1/disp?b_action=cognosViewer&ui.action=$($uiAction)&ui.object=$($camid)&&run.outputFormat=$($fileformat)&run.prompt=false$($reportparams)&cv.responseFormat=data"
} elseif ($uiAction -match "view") { #view a saved version of the report data
    $url = "$($baseURL)/$($cWebDir)/bi/v1/disp?b_action=cognosViewer&ui.action=$($uiAction)&ui.object=defaultOutput($($camid))&run.prompt=false$($reportparams)&cv.responseFormat=data"
}

if ($uiAction -notmatch "run" -and $uiAction -notmatch "view") {
    throw "Invalid uiAction option: use only 'view' or 'run'"
    Send-Email("[Failure][Invalid UI Action]")
    exit 2 #option not implemented
}

if ($showURL) {
    Write-Host "Verbose: Request sent to:", $([System.Web.HttpUtility]::UrlDecode("$url"))
}
#Write-Host $url
#Write-Host $([System.Web.HttpUtility]::UrlDecode($url))

trap { #general trap for errors
    $trap = $_
    Write-Host $trap.Exception.Message
    Send-Email("[Failure][Generic]",$($trap.Exception.Message))
    exit 9
}

if(!(Split-Path -parent $savepath) -or !(Test-Path -pathType Container (Split-Path -parent $savepath))) {
  $savepath = Join-Path $pwd (Split-Path -leaf $savepath)
}

$FileExists = Test-Path $fullfilepath
If ($FileExists -eq $True) {
    #replace datetime for if-modified-since header from existing file
    $filetimestamp = (Get-Item $fullfilepath).LastWriteTime
}

#submit login.
Write-Host "Attempting authentication..." -ForegroundColor Yellow
$response1 = Invoke-WebRequest -Uri "$($baseURL)/$($cWebDir)/bi/v1/login" -SessionVariable session -Method 'GET' -Credential $creds #-ErrorAction Ignore -SkipHttpErrorCheck

#switch to site.
Write-Host "Attempting switch into $dsnname..." -ForegroundColor Yellow
$response2 = Invoke-WebRequest -Uri "$($baseURL)/$($cWebDir)/bi/v1/login" -WebSession $session `
-Method "POST" `
-ContentType "application/json; charset=UTF-8" `
-Body "{`"parameters`":[{`"name`":`"h_CAM_action`",`"value`":`"logonAs`"},{`"name`":`"CAMNamespace`",`"value`":`"$camName`"},{`"name`":`"$dsnparam`",`"value`":`"$dsnname`"}]}"

$response3 = Invoke-WebRequest -Uri $url -WebSession $session
$HTMLDataString = $response3.Content

Write-Host("Downloaded HTML to retrieve report url.") -ForeGroundColor Yellow

$regex = [regex]"var sURL = '(.*?)'"

try { #Attempt to convert response to XML. If true the report is still processing and sURL hasn't been returned yet.
    if ([xml]$HTMLDataString) { $xmlresponse = $True }
} catch {
    $xmlresponse = $False
}

    if ($xmlresponse) {
        $reportjob = $($HTMLDataString).Replace('<xml><state>','').Replace('</state></xml>','').Replace('&quot;','"').Replace('\&','\\&') | ConvertFrom-Json
        
        if ($reportjob.m_sStatus -eq "working") {
            $trycount = 0
            $maxtry = 6
            do {

                $trycount++
                Write-Host("Report job status is [$($reportjob.m_sStatus)]. Waiting $($reportwait) seconds before check $($trycount) of $($maxtry)...") -ForeGroundColor Yellow
                Start-Sleep -Seconds $reportwait
                
                #take all fields returned from the XML and rebuild the string encoded.
                $poststring = 'b_action=cognosViewer&'
                $poststring += 'cv.actionState' + '=' + [System.Web.HttpUtility]::UrlEncode($reportjob.m_sActionState) + '&'
                $poststring += 'executionParameters' + '=' + [System.Web.HttpUtility]::UrlEncode($reportjob.m_sParameters) + '&'
                $poststring += 'm_tracking' + '=' + [System.Web.HttpUtility]::UrlEncode($reportjob.m_sTracking) + '&'
                $poststring += 'ui.cafcontextid' + '=' + [System.Web.HttpUtility]::UrlEncode($reportjob.m_sCAFContext) + '&'
                $poststring += 'ui.conversation' + '=' + [System.Web.HttpUtility]::UrlEncode($reportjob.m_sConversation) + '&'
                $poststring += 'ui.object' + '=' + "$camid" + '&'
                $poststring += 'ui.objectClass=report&'
                $poststring += 'ui.primaryAction=run&'
                $poststring += 'ui.action=wait&'
                $poststring += 'cv.responseFormat=data&'
                
                $response3 = Invoke-WebRequest -Uri "$($baseURL)/$($cWebDir)/bi/v1/disp" -WebSession $session -Method 'POST' -Body $poststring -ContentType "application/x-www-form-urlencoded"

                $HTMLDataString = $response3.Content
                Write-Host($HTMLDataString)
                
                try {
                    if ([xml]$HTMLDataString) { $xmlresponse = $True }
                } catch {
                    $xmlresponse = $False
                }

                if ($xmlresponse) {
                    $reportjob = $($HTMLDataString).Replace('<xml><state>','').Replace('</state></xml>','').Replace('&quot;','''').Replace('\&','\\&') | ConvertFrom-Json
                }
            } until (($HTMLDataString -match $regex) -or ($trycount -ge $maxtry))
        }
    }

    
    if ($HTMLDataString -notmatch $regex) {
        if ($HTMLDataString -match [regex]"RSV-CM-0005") { #Content Manager did not return an object for the requested search path
            Write-Host "Found 'RSV-CM-0005': File not found at the requested path. Try with or without specifying -ReportStudio?"
            exit 1
        }
        
        Write-Host($HTMLDataString) -ForeGroundColor Gray
        throw "'var sURL' not found"
        Send-Email("[Failure][sURL]")
        exit 3 #'var sURL' not found
    }

#Move Previous File
$PrevFileExists = Test-Path $fullfilepath
If ($PrevFileExists -eq $True) {
    $PrevOldFileExists = Test-Path ($fullfilepath + ".old")
    If ($PrevOldFileExists -eq $True) {
        Write-Host("Deleting old $report...") -ForeGroundColor Yellow
        Remove-Item -Path ($fullfilepath + ".old")
    }
    Write-Host("Renaming old $report...") -ForeGroundColor Yellow
    Rename-Item -Path $fullfilepath -newname ($fullfilepath + ".old")
}

$urlMatch = $regex.Matches($response3.Content)
$fileURLString = $urlMatch[0].Value.Replace("var sURL = '", "").Replace("'", "")

Invoke-WebRequest -Uri "$($baseURL)$($fileURLString)" -WebSession $session -OutFile $fullfilepath

# check file for proper format if csv
if ($extension -eq "csv") {
    $FileExists = Test-Path $fullfilepath
    if ($FileExists -eq $False) {
        Write-Host("Does not exist:" + $fullfilepath)
        Send-Email("[Failure][Output]","CSV Did not download to expected path.")
        exit 13 #CSV file didn't download to expected path
    }
    
    try {
        $filecontents = Import-CSV $fullfilepath

        $headercount = ($filecontents | Get-Member | Where-Object { $PSItem.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name | Measure-Object).Count
        if ($headercount -gt 1) {
            Write-Host("Passed CSV header check with $headercount headers...") -ForeGroundColor Yellow
        } else {
            Write-Host("Failed CSV header check with only $headercount headers...") -ForeGroundColor Yellow
            Reset-DownloadedFile($fullfilepath)
            Send-Email("[Failure][Verify]","Only $headercount header found in CSV.")
            exit 1
        }

        $linecount = ($filecontents | Measure-Object).Count
        if ($linecount -ge $requiredlinecount) { #Think schools.csv for smaller districts with only 3 campuses.
            Write-Host("Passed CSV line count with $linecount lines...") -ForeGroundColor Yellow
        } else {
            Write-Host("Failed CSV line count with only $linecount lines...") -ForeGroundColor Yellow
            Reset-DownloadedFile($fullfilepath)
            Send-Email("[Failure][Verify]","Only $linecount lines found in CSV.")
            exit 1
        }

    } catch {
        Send-Email("[Failure][Verify]")
        Reset-DownloadedFile($fullfilepath)
        exit 12 #reverted file format
    }
}

#need a valid exit here so this script can be put into a loop in case a file fails to download on first try
exit