Param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Give the name of the report you want to download.")][string]$report,
    [parameter(Position=1,Mandatory=$false,HelpMessage="Give a specific folder to download the report into.")][string]$savepath="C:\scripts", #--- VARIABLE --- change to a path you want to save files
    [parameter(Position=2,Mandatory=$false,HelpMessage="Extension to save onto the report name.")][string]$extension="CSV", #--- VARIABLE --- file extension to save data in csv or xlsx
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
    [parameter(Mandatory=$false)][switch]$ShowReportDetails, #Print report details to terminal.
    [parameter(Mandatory=$false)][switch]$SkipDownloadingFile, #Do not actually download the file.
    [parameter(Mandatory=$false)][switch]$dev, #print the url used to download report.
    [parameter(Mandatory=$false)][switch]$TeamContent #Report is in the Team Content folder.
)

Add-Type -AssemblyName System.Web

# The above parameters can be called directly from powershell switches
# In Cognos, do the following:
# 1. Setup a report with specific name (best without spaces like MyReportName) to run scheduled to save with which format you want then schedule this script to download it.
# 2. You will need to determine the DSN (database name) for your district
#    The database name is typically YOURSCHOOLNAMEsms (eSchool Plus) or YOURSCHOOLNAMEfms (eFinance)
#    Click the Login Avatar in the top right and click Sign In. Select esp (eSchool Plus) or efp (eFinance) then find your school database in the list.
# On computer to download data:
# 1. Create folder to store script and password data for script (default of C:\scripts)
# 2. Run from command line or batch script:
#    powershell.exe -executionpolicy bypass -file C:\Scripts\CognosDownload.ps1 -username 0000username -report MyReportName -cognosfolder "subfolder" -savepath "c:\scripts\downloads" -espdns schoolsms 
# 
# For eFinance: (coming soon)

#Example for the Team Content folder:
#.\CognosDownload.ps1 -username 0403cmillsap -report activities -cognosfolder "_Share Temporarily Between Districts/Gentry/automation" -espdsn gentrysms -TeamContent
#https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/wsil/path/Team%20Content%2FStudent%20Management%20System%2F_Share%20Temporarily%20Between%20Districts%2FGentry%2Fautomation
#/content/folder[@name='Student Management System']/folder[@name='_Share Temporarily Between Districts']/folder[@name='Gentry']/folder[@name='automation']/query[@name='activities']
#https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/atom/path/Team Content/Student Management System/_Share Temporarily Between Districts/Gentry/automation/activities
#CAMID("esp:a:0403cmillsap")/folder[@name='My Folders']/folder[@name='automation']/query[@name='activities']

# When the password expires, just delete the specific file (c:\scripts\apscnpw.txt) and run the script to re-create.

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
# 2020-12-01: Craig Millsap: Migrated over to using Invoke-RestMethod and better validation. Huge thanks to Jon Penn for the API documentation.

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
        Write-Host -NoNewline "Deleting old $report..." -ForeGroundColor Yellow
        Remove-Item -Path $fullfilepath -Force -ErrorAction SilentlyContinue
        Rename-Item -Path ($fullfilepath + ".old") -newname ($fullfilepath)
    }
    Write-Host "Reversing old $($report)." -ForeGroundColor Red
}

# server location for Cognos
if ($dev) {
    $baseURL = "https://dev.adecognos.arkansas.gov"
} else {
    $baseURL = "https://adecognos.arkansas.gov"
}

If ($eFinance) {
    $camName = "efp"    #efp for eFinance
    $dsnparam = "spi_db_name"
    $dsnname = $efpdsn
    $camid = "CAMID(""efp_x003Aa_x003A$($efpuser)"")"
} else {
    $camName = "esp"    #esp for eSchool
    $dsnparam = "dsn"
    $dsnname = $espdsn
    $camid = "CAMID(""esp_x003Aa_x003A$($username)"")"
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
    "pdf" { $fileformat = "PDF" }
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

if(!(Split-Path -parent $savepath) -or !(Test-Path -pathType Container (Split-Path -parent $savepath))) {
  $savepath = Join-Path $pwd (Split-Path -leaf $savepath)
}

$FileExists = Test-Path $fullfilepath
If ($FileExists -eq $True) {
    #replace datetime for if-modified-since header from existing file
    $filetimestamp = (Get-Item $fullfilepath).LastWriteTime
}

#submit login.
try {
    Write-Host -NoNewline "Attempting authentication... " -ForegroundColor Yellow
    $response1 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/login" -SessionVariable session -Method 'GET' -Credential $creds #-ErrorAction Ignore -SkipHttpErrorCheck
    Write-Host "Success." -ForegroundColor Yellow
} catch {
    Write-Host "Unable to authenticate." -ForegroundColor Red
    exit(1)
}

#switch to site.
try {
    Write-Host -NoNewline "Attempting switch into $dsnname... " -ForegroundColor Yellow
    $response2 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/login" -WebSession $session `
    -Method "POST" `
    -ContentType "application/json; charset=UTF-8" `
    -Body "{`"parameters`":[{`"name`":`"h_CAM_action`",`"value`":`"logonAs`"},{`"name`":`"CAMNamespace`",`"value`":`"$camName`"},{`"name`":`"$dsnparam`",`"value`":`"$dsnname`"}]}"
    Write-Host "Success." -ForegroundColor Yellow
} catch {
    Write-Host "Unable to switch into $dsnname." -ForegroundColor Red
    exit(2)
}

#No subfolder specified.
if ($cognosfolder -eq "My Folders") {
    #$cognosfolder = ([System.Web.HttpUtility]::UrlEncode("My Folders")).Replace('+','%20')
    $cognosfolder = "$($camid)/My Folders".Replace(' ','%20')
} elseif ($TeamContent) {
    $cognosfolder = "Team Content/Student Management System/$($cognosfolder)".Replace(' ','%20')
} else {
    #$cognosfolder = ([System.Web.HttpUtility]::UrlEncode("My Folders/$($cognosfolder)")).Replace('+','%20')
    $cognosfolder = "$($camid)/My Folders/$($cognosfolder)".Replace(' ','%20')
}

#Get the Atom feed
try {
    Write-Host -NoNewline "Attempting to retrieve report details for $($report)... " -ForegroundColor Yellow
    $response3 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/atom/path/$($cognosfolder)/$($report)" -WebSession $session
    $reportDetails = $response3.feed
    $reportID = $reportDetails.entry.storeID
    Write-Host "Success." -ForegroundColor Yellow
    
} catch {
    Write-Host "Unable to retrieve report details. Please check the supplied report name and cognosfolder." -ForegroundColor Red
    exit(3)
}

#Get the possible outputformats.
try {
    Write-Host -NoNewline "Retrieving possible formats... " -ForegroundColor Yellow
    $response4 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/outputFormats/path/$($cognosfolder)/$($report)" -WebSession $session
    Write-Host "Success." -ForegroundColor Yellow

    if ($response4.GetOutputFormatsResponse.supportedFormats.outputFormatName) {
        Write-Host " - $($report) ($($reportID)) can be exported in the following formats:" $($($response4.GetOutputFormatsResponse.supportedFormats.outputFormatName) -join ',') -ForegroundColor Yellow
    
        #This is case sensitive. So we need to retrieve the value from the response and match to the possible incorrect case provided to script.
        if ($($response4.GetOutputFormatsResponse.supportedFormats.outputFormatName) -contains $fileformat) {
            $possibleFormats = $($response4.GetOutputFormatsResponse.supportedFormats.outputFormatName)
            $validExtension = $possibleFormats[$($possibleFormats.ToLower().IndexOf($fileformat.ToLower()))]
        } else {
            Write-Host "You have requested an invalid extension type for this report."
            Throw "Invalid extension requested."
        }
    } else {
        Write-Host "Failed to retrieve output formats for the supplied report." -ForegroundColor Red
    }

} catch {
    Write-Host "Failed to retrieve output formats for the supplied report." -ForegroundColor Red
    exit(4)
}

#Print Additional Details to Terminal
if ($ShowReportDetails) {
    $details = $reportDetails | Select-Object -Property title,owner,ownerEmail,location,id
    $details.id = $reportID
    $details | Format-List
}

if (-Not($SkipDownloadingFile)) {
    Try {
        
        #Move Previous File
        $PrevFileExists = Test-Path $fullfilepath
        If ($PrevFileExists -eq $True) {
            $PrevOldFileExists = Test-Path ($fullfilepath + ".old")
            If ($PrevOldFileExists -eq $True) {
                Write-Host("Deleting old $report...") -ForeGroundColor Yellow
                Remove-Item -Path ($fullfilepath + ".old")
            }
            try {
                Write-Host -NoNewline "Renaming old $report... " -ForeGroundColor Yellow
                Rename-Item -Path $fullfilepath -newname ($fullfilepath + ".old")
                Write-Host "Success." -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to rename old report." -ForegroundColor Red
            }
        }

        Write-Host -NoNewline "Downloading Report to ""$($fullfilepath)""... " -ForegroundColor Yellow

        $downloadURL = "$($baseURL)/ibmcognos/bi/v1/disp/rds/outputFormat/path/$($cognosfolder)/$($report)/$($validExtension)?v=3"
        
        #https://www.ibm.com/support/knowledgecenter/SSEP7J_11.1.0/com.ibm.swg.ba.cognos.ca_dg_cms.doc/c_dg_raas_run_rep_prmpt.html#dg_raas_run_rep_prmpt
        #I think this should be a path as well to the xmlData so you can save it to a text file and pull in when needed to run.
        #Maybe if the prompts return a Test-Path $True then import and use the xmlData field instead. This should allow for more complex prompts.
        if ($reportparams -ne "") {
            $downloadURL = $downloadURL + '&' + $reportparams
        }
        
        $response6 = Invoke-RestMethod -Uri $downloadURL -WebSession $session -OutFile $fullfilepath

        Write-Host "Success." -ForegroundColor Yellow
    } catch {
        Write-Host "Failed to download file." -ForegroundColor Red
        exit(6)
    }
} else {
    #Just showing report details. No reason to continue.
    Write-Host "Skip downloading file specified. Exiting..." -ForegroundColor Yellow
    exit(0)
}

#Verify we didn't download an error page.
try {
    if (([xml](Get-Content $fullfilepath)).error.message) {

        $errorMessage = [xml](Get-Content $fullfilepath)
        Write-Host "Error detected in downloaded file. $($errorMessage.error.message)" -ForegroundColor Red

        if ($errorMessage.error.promptID) {
            $True
            #$a = Invoke-WebRequest -Uri "https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/reportPrompts/report/$(reportID)" -WebSession $session -SkipHttpErrorCheck -MaximumRedirection 15 -UseBasicParsing -ErrorAction SilentlyContinue -DisableKeepAlive
            #$a = Invoke-WebRequest -Uri "https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/sessionOutput/conversationID/$($errorMessage.error.PromptID)" -WebSession $session
            #$a = Invoke-RestMethod -Uri "$(baseURL)/ibmcognos/bi/v1/disp/rds/promptAnswers/conversationID/$($errorMessage.error.PromptID)" -WebSession $session
            #https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/sessionOutput/conversationID/iF684711856D041519349E983F76D7E79
            #$urlPath = ([System.Uri]($errorMessage.error.url)).PathAndQuery
            #$a = Invoke-RestMethod -Uri "$($baseURL)$($urlPath)" -WebSession $session
        }

        Move-Item -Path "$fullfilepath" -Destination "$($fullfilepath).error" -Force
        Reset-DownloadedFile($fullfilepath)
        exit(6)
    }
} catch {

}

# check file for proper format if csv
if ($extension -eq "csv") {
    $FileExists = Test-Path $fullfilepath
    if ($FileExists -eq $False) {
        Write-Host("Does not exist:" + $fullfilepath)
        Send-Email("[Failure][Output]","CSV Did not download to expected path.")
        exit(7) #CSV file didn't download to expected path
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
            exit(8)
        }

        $linecount = ($filecontents | Measure-Object).Count
        if ($linecount -ge $requiredlinecount) { #Think schools.csv for smaller districts with only 3 campuses.
            Write-Host("Passed CSV line count with $linecount lines...") -ForeGroundColor Yellow
        } else {
            Write-Host("Failed CSV line count with only $linecount lines...") -ForeGroundColor Yellow
            Reset-DownloadedFile($fullfilepath)
            Send-Email("[Failure][Verify]","Only $linecount lines found in CSV.")
            exit(9)
        }

    } catch {
        Send-Email("[Failure][Verify]")
        Reset-DownloadedFile($fullfilepath)
        exit (13) #General Verification Failure
    }
}

#need a valid exit here so this script can be put into a loop in case a file fails to download on first try
exit(0)