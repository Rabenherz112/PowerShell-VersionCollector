## WebScraper for Patch Management 
# Author: Jeremy Stephan
# Version: 1.0
# Date: 2022-12-17
# Description: This script get the latest Version for all used Software and save it in a HTML File and XML File
## 
$DATE = Get-Date -Format "yyyy-MM-dd"
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"; $MONTH = Get-Date -Format "MMMM"
$MONTH1 = Get-Date -Format "MM"
$YEAR = Get-Date -Format "yyyy"

# Define Script settings
$DEBUG = $false
$SEND_MAIL = $false

# Define file names
$OUTPUT_HTML_NAME = "LatestSoftwareVersions-$DATE.html"
$OUTPUT_XML_NAME = "LatestSoftwareVersions-$DATE.xml"
$OUTPUT_LOG_NAME = "LatestSoftwareVersions-$DATE.csv"

# Define file paths
$OUTPUT_HTML_PATH = "$PsScriptRoot\Output"
$OUTPUT_XML_PATH = "$PsScriptRoot\Output"
$OUTPUT_LOG_PATH = "$PsScriptRoot\Logs"

# Mail settings
$MAIL_FROM = "dummy@email.com"
$MAIL_TO = @("dummy@email.com", "dummy@email.com")
$MAIL_SUBJECT = "Patchmgmt: Monthly Version Update ($DATE)"
$MAIL_BODY = Get-Content -Path "D:\GIT\PowerShell-VersionCollector\Templates\MailNotifyTemplate.html" -Raw
$MAIL_SERVER = "mail.server.com"
$MAIL_PORT = 25
$MAIL_USER = $MAIL_FROM
$MAIL_PASSWORD = "PASSWORD"
$MAIL_USE_CREDENTIALS = $false
$MAIL_USE_SSL = $false
$MAIL_MANUALLY_UPDATE = @"
- Other Software 1 <br>
- Other Software 2 <br>
"@

## Define all URLs
# Microsoft Windows Update History
$URL_WC11_22H2 = "https://support.microsoft.com/en-us/topic/windows-11-version-22h2-update-history-ec4229c3-9c5f-4e75-9d6d-9025ab70fcce"
$URL_WC11_21H2 = "https://support.microsoft.com/en-us/topic/windows-11-version-21h2-update-history-a19cd327-b57f-44b9-84e0-26ced7109ba9"
$URL_WC10_22H2 = "https://support.microsoft.com/en-us/topic/windows-10-update-history-8127c2c6-6edf-4fdf-8b9f-0f7be1ef3562"
$URL_WC10_21H2 = "https://support.microsoft.com/en-us/topic/windows-10-update-history-857b8ccb-71e4-49e5-b3f6-7073197d98fb"
$URL_WC10_21H1 = "https://support.microsoft.com/en-us/topic/windows-10-update-history-1b6aac92-bf01-42b5-b158-f80c6d93eb11"
$URL_WC10_20H2 = "https://support.microsoft.com/en-us/topic/windows-10-update-history-7dd3071a-3906-fa2c-c342-f7f86728a6e3"
$URL_WC10_1909 = "https://support.microsoft.com/en-us/topic/windows-10-update-history-53c270dc-954f-41f7-7ced-488578904dfe"
$URL_WC81 = "https://support.microsoft.com/en-us/topic/windows-8-1-and-windows-server-2012-r2-update-history-47d81dd2-6804-b6ae-4112-20089467c7a6"
$URL_WS_2012 = "https://support.microsoft.com/en-us/topic/windows-server-2012-update-history-abfb9afd-2ebf-1c19-4224-ad86f8741edd"
$URL_WS_2012R = "https://support.microsoft.com/en-us/topic/windows-8-1-and-windows-server-2012-r2-update-history-47d81dd2-6804-b6ae-4112-20089467c7a6"
$URL_WS_2016 = "https://support.microsoft.com/en-us/topic/windows-10-and-windows-server-2016-update-history-4acfbc84-a290-1b54-536a-1c0430e9f3fd"
$URL_WS_2019 = "https://support.microsoft.com/en-us/topic/windows-10-and-windows-server-2019-update-history-725fc2e1-4443-6831-a5ca-51ff5cbcb059"
$URL_WS_2022 = "https://support.microsoft.com/en-gb/topic/windows-server-2022-update-history-e1caa597-00c5-4ab9-9f3e-8212fe80b2ee"
# Microsoft Office Update History
$URL_OFFICE_365 = "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
$URL_OFFICE_2021 = "https://learn.microsoft.com/en-us/officeupdates/update-history-office-2021"
$URL_OFFICE_2019 = "https://learn.microsoft.com/en-us/officeupdates/update-history-office-2019"
$URL_OFFICE_2016 = "https://learn.microsoft.com/en-us/officeupdates/update-history-office-2019"
# Other Microsoft Products
$URL_AZURE_DATASTUDIO = "https://learn.microsoft.com/en-us/sql/azure-data-studio/release-notes-azure-data-studio"
# Other Products
$URL_DOCKER_DESKTOP = "https://docs.docker.com/desktop/release-notes"
$URL_CISCO_ANNYCONNECT = "https://www.cisco.com/c/en/us/td/docs/security/vpn_client/anyconnect/Cisco-Secure-Client-5/release/notes/release-notes-cisco-secure-client-5-0.html"
$URL_DEVOLUTIONS_RDM = "https://devolutions.net/remote-desktop-manager/release-notes"
$URL_GOTO_MEETING = "https://support.goto.com/meeting/help/whats-new-in-goto-meeting"
$URL_OBS_STUDIO = "https://obsproject.com/osx_update/stable/notes.html"
$URL_POWERSHELL7 = "https://github.com/PowerShell/powershell/releases"
$URL_SENCHA_CMD = "https://www.sencha.com/products/extjs/cmd-download/"
$URL_SWAGGER_UI = "https://github.com/swagger-api/swagger-ui/releases"

# Microsoft Windows Data Collection Config
$INPUTDATA_MS_WINDOWS = @(
    @{Name = "Windows 11 22H2"; URL = $URL_WC11_22H2; XMLName = "WindowsVerClient11-22H2" },
    @{Name = "Windows 11 21H2"; URL = $URL_WC11_21H2; XMLName = "WindowsVerClient11-21H2" },
    @{Name = "Windows 10 22H2"; URL = $URL_WC10_22H2; XMLName = "WindowsVerClient22H2" },
    @{Name = "Windows 10 21H2"; URL = $URL_WC10_21H2; XMLName = "WindowsVerClient21H2" },
    @{Name = "Windows 10 21H1"; URL = $URL_WC10_21H1; XMLName = "WindowsVerClient21H1" },
    @{Name = "Windows 10 20H2"; URL = $URL_WC10_20H2; XMLName = "WindowsVerClient20H2" },
    @{Name = "Windows 10 1909"; URL = $URL_WC10_1909; XMLName = "WindowsVerClient1909" },
    @{Name = "Windows 8.1"; URL = $URL_WC81; XMLName = "WindowsVerClient81" }
    @{Name = "Windows Server 2022"; URL = $URL_WS_2022; XMLName = "WindowsVerServer2022" },
    @{Name = "Windows Server 2019"; URL = $URL_WS_2019; XMLName = "WindowsVerServer2019" },
    @{Name = "Windows Server 2016"; URL = $URL_WS_2016; XMLName = "WindowsVerServer2016" },
    @{Name = "Windows Server 2012 R2"; URL = $URL_WS_2012R; XMLName = "WindowsVerServer2012R2" },
    @{Name = "Windows Server 2012"; URL = $URL_WS_2012; XMLName = "WindowsVerServer2012" }
)
# Microsoft Office Data Collection Config
$INPUTDATA_MS_OFFICE = @(
    @{Name = "Office 365"; URL = $URL_OFFICE_365; XMLName = "AppVerMSOffice365" },
    @{Name = "Office 2021"; URL = $URL_OFFICE_2021; XMLName = "AppVerMSOffice2021" },
    @{Name = "Office 2019"; URL = $URL_OFFICE_2019; XMLName = "AppVerMSOffice2019" },
    @{Name = "Office 2016"; URL = $URL_OFFICE_2016; XMLName = "AppVerMSOffice2016" }
)

# Logging function that will be used to log all actions
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information', 'Warning', 'Error', 'Debug', 'Data')]
        [string]$Severity = 'Information'
    )

    # Write to console
    switch ($Severity) {
        'Debug' {
            if ($DEBUG) {
                Write-Host -NoNewline -ForegroundColor DarkGray -Object " [DBG] "
            }
            else {
                return
            } 
        }
        'Information' { Write-Host -NoNewline -ForegroundColor Green -Object " [INFO] " }
        'Data' { Write-Host -NoNewline -ForegroundColor Cyan -Object " [DATA] " }
        'Warning' { Write-Host -NoNewline -ForegroundColor Yellow -Object " [WARN] " }
        'Error' { Write-Host -NoNewline -ForegroundColor Red -Object " [ERR] " }
    }
    Write-Host -ForegroundColor White -Object "$Message"
 
    # Write to log file
    [pscustomobject]@{
        Time     = (Get-Date -f g)
        Message  = $Message
        Severity = $Severity
    } | Export-Csv -Path "$OUTPUT_LOG_PATH\$OUTPUT_LOG_NAME" -Append -NoTypeInformation
}
Write-Log -Message "Initialized Logging function" -Severity Debug

## Script Safety Checks
# Check if configuration variables are set
if (!$OUTPUT_HTML_PATH -or !$OUTPUT_XML_PATH -or !$OUTPUT_LOG_PATH -or !$OUTPUT_HTML_NAME -or !$OUTPUT_XML_NAME -or !$OUTPUT_LOG_NAME) {
    Write-Log -Message "Please set all configuration variables" -Severity Error
    exit 1
}

# Try to create Output Directories if they do not exist
if (!(Test-Path $OUTPUT_HTML_PATH)) {
    Write-Log -Message "Creating HTML Output Directory" -Severity Debug
    New-Item -ItemType Directory -Path $OUTPUT_HTML_PATH
}
if (!(Test-Path $OUTPUT_XML_PATH)) {
    Write-Log -Message "Creating XML Output Directory" -Severity Debug
    New-Item -ItemType Directory -Path $OUTPUT_XML_PATH
}
if (!(Test-Path $OUTPUT_LOG_PATH)) {
    Write-Log -Message "Creating Log Output Directory" -Severity Debug
    New-Item -ItemType Directory -Path $OUTPUT_LOG_PATH
}

# Check if all Paths exist
if (!(Test-Path $OUTPUT_HTML_PATH) -or !(Test-Path $OUTPUT_XML_PATH) -or !(Test-Path $OUTPUT_LOG_PATH)) {
    Write-Log -Message "One or more output paths do not exist" -Severity Error
    exit 1
}
Write-Log -Message "All configuration variables are set" -Severity Debug

# Check if Output Files already exist if so warn User and ask if Files should be overwritten
if (Test-Path "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME") {
    Write-Log -Message "HTML File already exists" -Severity Warning
    $OverwriteHTML = Read-Host -Prompt "Do you want to overwrite the HTML File? (y/n)"
    if ($OverwriteHTML -eq "y") {
        Write-Log -Message "Overwriting HTML File" -Severity Debug
    }
    else {
        Write-Log -Message "Stopping Script to prevent deletion of HTML File" -Severity Warning
        Write-Log -Message "HTML is located at $OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME" -Severity Information
        exit 1
    }
}
if (Test-Path "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME") {
    Write-Log -Message "XML File already exists" -Severity Warning
    $OverwriteXML = Read-Host -Prompt "Do you want to overwrite the XML File? (y/n)"
    if ($OverwriteXML -eq "y") {
        Write-Log -Message "Overwriting XML File" -Severity Debug
    }
    else {
        Write-Log -Message "Stopping Script to prevent deletion of XML File" -Severity Warning
        Write-Log -Message "XML is located at $OUTPUT_XML_PATH\$OUTPUT_XML_NAME" -Severity Information
        exit 1
    }
}
Write-Log -Message "Checked all File Locations" -Severity Debug

# Check if PowerShell Version is 5.X or higher is used for running the script (PS 7 crashes for some reason)
if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Log -Message "This Script need to be run in PowerShell Version 5.1 or a higher Minor Version" -Severity Error
    Write-Log -Message "Current PowerShell Version is $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)" -Severity Information
    exit 1
}
Write-Log -Message "PowerShell Version is 5.1 or higher" -Severity Debug
# Check if PowerShell 7 is installed (for dependency)
$PS7Exists = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\PowerShellCore\InstalledVersions | Get-ItemProperty -Name SemanticVersion
$PS7Exists = $PS7Exists.SemanticVersion.Split(".")[0]
if ($PS7Exists -ne 7) {
    Write-Log -Message "This Script requires PowerShell 7 to be installed" -Severity Error
    exit 1
}
Write-Log -Message "PowerShell 7 is installed" -Severity Debug

# Check if Internet Connection is available
if (!(Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet)) {
    Write-Log -Message "This Script requires an Internet Connection" -Severity Error
}
Write-Log -Message "Internet Connection is available" -Severity Debug

## File Creation and Initialization
# Function add Application Names and Versions to HTML File
$ACTIVE_HTML = @"
<!DOCTYPE html>
<html>
    <head>
        <title>Patch Management Software Versions</title>
        <style>
            body {
                font-family: Arial, Helvetica, sans-serif;
                font-size: 12px;
                color: #000000;
                background-color: #FFFFFF;
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th {
                padding: 8px;
                text-align: left;
                background-color: #4CAF50;
                color: white;
            }
            tr:nth-child(even) {
                background-color: #f2f2f2;
            }
            td {
                padding: 8px;
            }
            .markcellred {
                background-color: #FF0000;
            }
            .markcellyellow {
                background-color: #FFFF00;
            }
            .footer {
                position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                color: #808080;
                text-align: center;
            }
        </style>
    </head>
    <body>
        <h1>Latest Software Versions</h1>
        <table>
            <tr>
                <th>Application</th>
                <th>Version</th>
            </tr>
"@
$ACTIVE_HTML | Out-File -FilePath "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME" -Encoding UTF8

function Add-HTML {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationName,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationVersion,

        [Parameter()]
        [ValidateSet('red', 'yellow')]
        [string]$ApplicationMarkup
    )

    $ACTIVE_HTML = Get-Content -Path "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME"
    switch ($ApplicationMarkup) {
        "red" { $ACTIVE_HTML += "<tr><td class='markcellred'>$ApplicationName</td><td class='markcellred'>$ApplicationVersion</td></tr>" }
        "yellow" { $ACTIVE_HTML += "<tr><td class='markcellyellow'>$ApplicationName</td><td class='markcellyellow'>$ApplicationVersion</td></tr>" }
        default { $ACTIVE_HTML += "<tr><td>$ApplicationName</td><td>$ApplicationVersion</td></tr>" }
    }
    $ACTIVE_HTML | Out-File -FilePath "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME" -Encoding UTF8
    Write-Log -Message "Added $ApplicationName with value $ApplicationVersion to HTML File" -Severity Debug
}
Write-Log -Message "Initialized HTML Converter" -Severity Debug

# Function add Application Names and Versions to XML File
$ACTIVE_XML = @"
<?xml version="1.0" encoding="utf-8"?>
<AdminArsenal.Export Code="PDQInventory" Name="PDQ Inventory" Version="19.3.360.0" MinimumVersion="5.0">
  <VariablesSettingsViewModel>
    <CustomVariables type="list">
"@
$ACTIVE_XML | Out-File -FilePath "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME" -Encoding UTF8

function Add-XML {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationName,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationVersion
    )

    $ACTIVE_XML = Get-Content -Path "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME"
    $ACTIVE_XML += 
    @"
      <CustomVariable>
        <Name>$ApplicationName</Name>
        <Value>$ApplicationVersion</Value>
      </CustomVariable>
"@
    $ACTIVE_XML | Out-File -FilePath "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME" -Encoding UTF8
    Write-Log -Message "Added $ApplicationName with value $ApplicationVersion to XML File" -Severity Debug
}
Write-Log -Message "Initialized XML Converter" -Severity Debug

## Other Functions
# Create inputbox for user input
function CustomUserInput {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Title,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$DefaultText
    )
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $UserInput = [Microsoft.VisualBasic.Interaction]::InputBox($Message, "Patchmgmt: $Title", $DefaultText)
    Write-Log -Message "User entered $UserInput" -Severity Debug
    return $UserInput
}
Write-Log -Message "Initialized UserInput Function" -Severity Debug
#$demo = CustomUserInput -Message "Please enter your name" -Title "Name" -DefaultText "John Doe"

### Start of Software Version Collection
Write-Log -Message "Starting Microsoft Windows Data Collection" -Severity Debug
foreach ($INPUTDATA_MS_WINDOWS_ENTRY in $INPUTDATA_MS_WINDOWS) {
    # Get latest $($INPUTDATAENTRY.Name)
    Write-Log -Message "Getting latest $($INPUTDATA_MS_WINDOWS_ENTRY.Name)" -Severity Information
    try {
        $HTML = Invoke-WebRequest -Uri $($INPUTDATA_MS_WINDOWS_ENTRY.URL) -UseBasicParsing
        $HTML2 = New-Object -Com "HTMLFile"
        [string]$HTMLBODY = $HTML.Content
        $HTML2.write([ref]$HTMLBODY)
        $HTMLINNER = $html2.getElementsByClassName("supLeftNavActiveCategory")[0].innerText
    }
    catch {
        Write-Log -Message "Could not get latest $($INPUTDATA_MS_WINDOWS_ENTRY.Name) via $($INPUTDATA_MS_WINDOWS_ENTRY.URL)" -Severity Error
    }
    ## OLD CODE
    #$VERSION_EXISTS = $HTMLINNER -match "$MONTH [0-9]{1,2}, $YEAR(.*)"
    #$VERSION_EXISTS_TEMP = $Matches[0] -match "KB[0-9]{6,7}"
    #$APPVERSION = $Matches[0]
    ## Check if Windows Version is EOL
    #$EOL = $HTMLINNER -match "End of service statement"
    $VERSION_EXISTS = $HTMLINNER -match ".*(KB[\d]*) .*"
    $APPVERSION = $Matches[1]
    # Check if Windows Version is EOL
    $EOL = $HTMLINNER -match "End of service statement"
    if (!$VERSION_EXISTS) {
        Write-Log -Message "Could not find latest $($INPUTDATA_MS_WINDOWS_ENTRY.Name) KB" -Severity Warning
        if ($EOL) {
            Write-Log -Message "$($INPUTDATA_MS_WINDOWS_ENTRY.Name) is EOL" -Severity Warning
            Add-HTML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.Name)" -ApplicationVersion "EOL" -ApplicationMarkup "red"
            Add-XML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.XMLName)" -ApplicationVersion "EOL"
        }
    }
    else {
        if ($EOL) {
            Write-Log -Message "Latest $($INPUTDATA_MS_WINDOWS_ENTRY.Name) is $APPVERSION (EOL)" -Severity Data
            Add-HTML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.Name)" -ApplicationVersion "EOL $APPVERSION" -ApplicationMarkup "yellow"
            Add-XML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.XMLName)" -ApplicationVersion "$APPVERSION"
        }
        else {
            Write-Log -Message "Latest $($INPUTDATA_MS_WINDOWS_ENTRY.Name) is $APPVERSION" -Severity Data
            Add-HTML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.Name)" -ApplicationVersion "$APPVERSION"
            Add-XML -ApplicationName "$($INPUTDATA_MS_WINDOWS_ENTRY.XMLName)" -ApplicationVersion "$APPVERSION"
        }
    }
}
Write-Log -Message "Finished Microsoft Windows Data Collection" -Severity Debug

### Start of Microsoft Office Data Collection
Write-Log -Message "Starting Microsoft Office Data Collection" -Severity Debug
foreach ($INPUTDATA_MS_OFFICE_ENTRY in $INPUTDATA_MS_OFFICE) {
    Write-Log -Message "Getting latest $($INPUTDATA_MS_OFFICE_ENTRY.Name)" -Severity Information
    try {
        $HTML = Invoke-WebRequest -Uri $($INPUTDATA_MS_OFFICE_ENTRY.URL) -UseBasicParsing
        $HTML2 = New-Object -Com "HTMLFile"
        [string]$HTMLBODY = $HTML.Content
        $HTML2.write([ref]$HTMLBODY)
        $HTMLINNER = $HTML2.getElementsByTagName("table")[0]
        $HTMLINNER = $HTMLINNER.getElementsByTagName("tr")[1]
        if ($INPUTDATA_MS_OFFICE_ENTRY.Name -eq "Office 365") {
            $HTMLINNER = $HTMLINNER.getElementsByTagName("td")[2]
            $HTMLINNER = $HTMLINNER.innerText
            Write-Log -Message "Office 365 Version is $HTMLINNER" -Severity Data
            Add-HTML -ApplicationName "$($INPUTDATA_MS_OFFICE_ENTRY.Name)" -ApplicationVersion "$HTMLINNER"
            Add-XML -ApplicationName "$($INPUTDATA_MS_OFFICE_ENTRY.XMLName)" -ApplicationVersion "$HTMLINNER"
        }
        else {
            $HTMLINNER = $HTMLINNER.getElementsByTagName("td")[1]
            $HTMLINNER = $HTMLINNER.innerText
            $HTMLINNER = $HTMLINNER -split "Build "
            $HTMLINNER = $HTMLINNER[1]
            $HTMLINNER = $HTMLINNER -split "\)"
            $HTMLINNER = $HTMLINNER[0]
            Write-Log -Message "$($INPUTDATA_MS_OFFICE_ENTRY.Name) Version is $HTMLINNER" -Severity Data
            Add-HTML -ApplicationName "$($INPUTDATA_MS_OFFICE_ENTRY.Name)" -ApplicationVersion "$HTMLINNER"
            Add-XML -ApplicationName "$($INPUTDATA_MS_OFFICE_ENTRY.XMLName)" -ApplicationVersion "$HTMLINNER"
        }
    }
    catch {
        Write-Log -Message "Could not get latest $($INPUTDATA_MS_OFFICE_ENTRY.Name) via $($INPUTDATA_MS_OFFICE_ENTRY.URL)" -Severity Error
    }
}
Write-Log -Message "Finished Microsoft Office Data Collection" -Severity Debug

# Start of other Microsoft Software Data Collection
Write-Log -Message "Starting other Microsoft Software Data Collection" -Severity Debug
try {
    Write-Log -Message "Getting latest Microsoft Azure Data Studio" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_AZURE_DATASTUDIO -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $HTMLINNER = $HTML2.getElementsByTagName("h4")[0]
    $HTMLINNER = $HTMLINNER.innerText
    # What's new in 1.42.0 - Get only the version number with regex
    $VERSION_EXISTS = $HTMLINNER -match ".*([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2})"
    $HTMLINNER = $Matches[1]
    Write-Log -Message "Microsoft Azure Data Studio Version is $HTMLINNER" -Severity Data
    Add-HTML -ApplicationName "Azure Data Studio" -ApplicationVersion "$HTMLINNER"
    Add-XML -ApplicationName "AppVerMSAzureDataStudio" -ApplicationVersion "$HTMLINNER"
}
catch {
    Write-Log -Message "Could not get latest Microsoft Azure Data Studio" -Severity Error
}
Write-Log -Message "Finished other Microsoft Software Data Collection" -Severity Debug

# Start of other Software Data Collection
Write-Log -Message "Starting other Software Data Collection" -Severity Debug
try {
    Write-Log -Message "Getting latest Docker Desktop" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_DOCKER_DESKTOP -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $HTMLINNER = $HTML2.getElementsByTagName("h2")[0]
    $HTMLINNER = $HTMLINNER.innerText
    Write-Log -Message "Docker Desktop Version is $HTMLINNER" -Severity Data
    Add-HTML -ApplicationName "Docker Desktop" -ApplicationVersion "$HTMLINNER"
    Add-XML -ApplicationName "AppVerDockerWindows" -ApplicationVersion "$HTMLINNER"
}
catch {
    Write-Log -Message "Could not get latest Docker Desktop" -Severity Error
}

try {
    Write-Log -Message "Getting latest Cisco AnyConnect" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_CISCO_ANNYCONNECT -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $VERSION_EXISTS_TEMP = $HTMLBODY -match "Cisco Secure Client (.*) New Features"
    $APPVERSION = $Matches[1]
    Write-Log -Message "Cisco AnyConnect Version is $APPVERSION" -Severity Data
    Add-HTML -ApplicationName "Cisco AnyConnect" -ApplicationVersion "$APPVERSION"
    Add-XML -ApplicationName "AppVerCiscoAnyConnect" -ApplicationVersion "$APPVERSION"
}
catch {
    Write-Log -Message "Could not get latest Cisco AnyConnect" -Severity Error
}

try {
    Write-Log -Message "Getting latest Remote Desktop Manager (Devolutions)" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_DEVOLUTIONS_RDM -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $HTMLINNER = $HTML2.getElementsByTagName("h4")[0]
    $HTMLINNER = $HTMLINNER.innerText
    $HTMLINNER = $HTMLINNER -split "Version "
    $HTMLINNER = $HTMLINNER[1]
    $HTMLINNER = $HTMLINNER -split " "
    $HTMLINNER = $HTMLINNER[0]
    Write-Log -Message "Remote Desktop Manager (Devolutions) Version is $HTMLINNER" -Severity Data
    Add-HTML -ApplicationName "Remote Desktop Manager (Devolutions)" -ApplicationVersion "$HTMLINNER"
    Add-XML -ApplicationName "AppVerDevolutionsRDM" -ApplicationVersion "$HTMLINNER"
}
catch {
    Write-Log -Message "Could not get latest Remote Desktop Manager (Devolutions)" -Severity Error
}

try {
    Write-Log -Message "Getting latest GoTo Meeting" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_GOTO_MEETING -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $VERSION_EXISTS_TEMP = $HTMLBODY -match "Desktop app \((.*)\)"
    $APPVERSION = $Matches[1]
    $APPVERSION = $APPVERSION -replace "v", ""
    $APPVERSION = $APPVERSION -replace "b", ""
    $APPVERSION = $APPVERSION -replace ", ", "."
    Write-Log -Message "GoTo Meeting Version is $APPVERSION" -Severity Data
    Add-HTML -ApplicationName "GoTo Meeting" -ApplicationVersion "$APPVERSION"
    Add-XML -ApplicationName "AppVerGoToMeeting" -ApplicationVersion "$APPVERSION"
}
catch {
    Write-Log -Message "Could not get latest GoTo Meeting" -Severity Error
}

try {
    Write-Log -Message "Getting latest OBS Studio" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_OBS_STUDIO -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $HTMLINNER = $HTML2.getElementsByTagName("h1")[0]
    $HTMLINNER = $HTMLINNER.innerText
    $HTMLINNER = $HTMLINNER -split " "
    $HTMLINNER = $HTMLINNER[2]
    Write-Log -Message "OBS Studio Version is $HTMLINNER" -Severity Data
    Add-HTML -ApplicationName "OBS Studio" -ApplicationVersion "$HTMLINNER"
    Add-XML -ApplicationName "AppVerOBS" -ApplicationVersion "$HTMLINNER"
}
catch {
    Write-Log -Message "Could not get latest OBS Studio" -Severity Error
}

try {
    Write-Log -Message "Getting latest PowerShell 7" -Severity Information
    $EXPERIMENTAL = $true
    $HTML = Invoke-WebRequest -Uri $URL_POWERSHELL7 -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $COUNT = 0
    do {
        # Check for all Links on the Page and check if it contains v(.*) Release
        $HTMLINNER = $HTML2.getElementsByTagName("a")[$COUNT]
        $HTMLINNER = $HTMLINNER.innerText
        $VERSION_EXISTS_TEMP = $HTMLINNER -match "v(.*) Release"
        if ($VERSION_EXISTS_TEMP) {
            $APPVERSION = $Matches[1]
            # Check if AppVersion does not any letters
            $VERSION_EXISTS_TEMP = $APPVERSION -match "[a-zA-Z]"
            if ($VERSION_EXISTS_TEMP) {
                # If it does, then it is an experimental version
                $EXPERIMENTAL = $true
            }
            else {
                # If it does not, then it is the latest version
                $EXPERIMENTAL = $false
            }
        }
        $COUNT += 1
    } while ($EXPERIMENTAL)
    Write-Log -Message "PowerShell 7 Version is $APPVERSION" -Severity Data
    Add-HTML -ApplicationName "PowerShell 7" -ApplicationVersion "$APPVERSION"
    Add-XML -ApplicationName "AppVerPowerShell7.0" -ApplicationVersion "$APPVERSION"
}
catch {
    Write-Log -Message "Could not get latest PowerShell 7" -Severity Error
}

try {
    Write-Log -Message "Getting lastest Sencha Cmd" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_SENCHA_CMD -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $HTMLINNER = $HTML2.getElementsByTagName("h2")[0]
    $HTMLINNER = $HTMLINNER.innerText
    $HTMLINNER = $HTMLINNER -split " "
    $HTMLINNER = $HTMLINNER[1]
    Write-Log -Message "Sencha Cmd Version is $HTMLINNER" -Severity Data
    Add-HTML -ApplicationName "Sencha Cmd" -ApplicationVersion "$HTMLINNER"
    Add-XML -ApplicationName "AppVerSenchaCmd" -ApplicationVersion "$HTMLINNER"
}
catch {
    Write-Log -Message "Could not get latest Sencha Cmd" -Severity Error
}

try {
    Write-Log -Message "Getting latest Swagger UI" -Severity Information
    $HTML = Invoke-WebRequest -Uri $URL_SWAGGER_UI -UseBasicParsing
    $HTML2 = New-Object -Com "HTMLFile"
    [string]$HTMLBODY = $HTML.Content
    $HTML2.write([ref]$HTMLBODY)
    $COUNT = 0
    $EXPERIMENTAL = $true
    do {
        # Check for all Links on the Page and check if it contains v(.*) Release
        $HTMLINNER = $HTML2.getElementsByTagName("a")[$COUNT]
        $HTMLINNER = $HTMLINNER.innerText
        $VERSION_EXISTS_TEMP = $HTMLINNER -match "v[0-9](.*)"
        if ($VERSION_EXISTS_TEMP) {
            $APPVERSION = $Matches[0]
            $APPVERSION = $APPVERSION -replace "v", ""
            $APPVERSION = $APPVERSION -replace " ", ""
            $VERSION_EXISTS_TEMP = $APPVERSION -match "[a-zA-Z]"
            if ($VERSION_EXISTS_TEMP) {
                # If it does, then it is an experimental version
                $EXPERIMENTAL = $true
            }
            else {
                # If it does not, then it is the latest version
                $EXPERIMENTAL = $false
            }
        }
        $COUNT += 1
    } while ($EXPERIMENTAL)
    Write-Log -Message "Swagger UI Version is $APPVERSION" -Severity Data
    Add-HTML -ApplicationName "Swagger UI" -ApplicationVersion "$APPVERSION"
    Add-XML -ApplicationName "AppVerSwaggerUI" -ApplicationVersion "$APPVERSION"
}
catch {
    Write-Log -Message "Could not get latest Swagger UI" -Severity Error
}
# Generated fake Data for testing
<#for ($i = 0; $i -lt 60; $i++) {
    $ApplicationName = "Application $i"
    $ApplicationVersion = "Version $i"
    Add-HTML -ApplicationName $ApplicationName -ApplicationVersion $ApplicationVersion
    Add-XML -ApplicationName $ApplicationName -ApplicationVersion $ApplicationVersion
}
#>

Write-Log -Message "Finished Getting Application Versions" -Severity Debug

# Finish HTML File
$ACTIVE_HTML = Get-Content -Path "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME"
$ACTIVE_HTML += @"
        </table>
        <div class='footer'>
            <p>Generated by $env:computername at $(Get-Date -f g)</p>
            <p>Script by <a href='https://jerst.net'>Jeremy Stephan</a></p>
        </div>
    </body>
</html>
"@
$ACTIVE_HTML | Out-File -FilePath "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME" -Encoding UTF8
Write-Log -Message "Finished HTML File" -Severity Debug

# Finish XML File
$ACTIVE_XML = Get-Content -Path "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME"
$ACTIVE_XML += @"
    </CustomVariables>
  </VariablesSettingsViewModel>
</AdminArsenal.Export>
"@
$ACTIVE_XML | Out-File -FilePath "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME" -Encoding UTF8
Write-Log -Message "Finished XML File" -Severity Debug

# Send Email
if ($SEND_MAIL -eq $true) {
    Write-Log -Message "Sending Email" -Severity Debug
    if ($MAIL_BODY -imatch "%YEAR%") {
        $MAIL_BODY = $MAIL_BODY -replace "%YEAR%", $YEAR
    }
    if ($MAIL_BODY -imatch "%MONTH%") {
        $MAIL_BODY = $MAIL_BODY -replace "%MONTH%", $MONTH1
    }
    if ($MAIL_BODY -imatch "%COMPUTERNAME%") {
        $MAIL_BODY = $MAIL_BODY -replace "%COMPUTERNAME%", $env:computername
    }
    if ($MAIL_BODY -imatch "%DATE%") {
        $MAIL_BODY = $MAIL_BODY -replace "%DATE%", $(Get-Date -f g)
    }
    if ($MAIL_BODY -imatch "%MANUAL_UPDATE%") {
        $MAIL_BODY = $MAIL_BODY -replace "%MANUAL_UPDATE%", $MAIL_MANUALLY_UPDATE
    }
    try {
        if ($MAIL_USE_SSL -and $MAIL_USE_CREDENTIALS) {
            Send-MailMessage -From $MAIL_FROM -To $MAIL_TO -Subject $MAIL_SUBJECT -Body $MAIL_BODY -BodyAsHtml -Encoding UTF8 -SmtpServer $MAIL_SERVER -UseSsl -Port $MAIL_PORT -Credential $MAIL_CREDENTIALS -Attachments "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME", "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME", "$OUTPUT_LOG_PATH\$OUTPUT_LOG_NAME"
        }
        elseif (!($MAIL_USE_SSL) -and $MAIL_USE_CREDENTIALS) {
            Send-MailMessage -From $MAIL_FROM -To $MAIL_TO -Subject $MAIL_SUBJECT -Body $MAIL_BODY -BodyAsHtml -Encoding UTF8 -SmtpServer $MAIL_SERVER -Port $MAIL_PORT -Credential $MAIL_CREDENTIALS -Attachments "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME", "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME", "$OUTPUT_LOG_PATH\$OUTPUT_LOG_NAME"
        }
        elseif ($MAIL_USE_SSL -and !($MAIL_USE_CREDENTIALS)) {
            Send-MailMessage -From $MAIL_FROM -To $MAIL_TO -Subject $MAIL_SUBJECT -Body $MAIL_BODY -BodyAsHtml -Encoding UTF8 -SmtpServer $MAIL_SERVER -UseSsl -Port $MAIL_PORT -Attachments "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME", "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME", "$OUTPUT_LOG_PATH\$OUTPUT_LOG_NAME"
        }
        else {
            Send-MailMessage -From $MAIL_FROM -To $MAIL_TO -Subject $MAIL_SUBJECT -Body $MAIL_BODY -BodyAsHtml -Encoding UTF8 -SmtpServer $MAIL_SERVER -Port $MAIL_PORT -Attachments "$OUTPUT_HTML_PATH\$OUTPUT_HTML_NAME", "$OUTPUT_XML_PATH\$OUTPUT_XML_NAME", "$OUTPUT_LOG_PATH\$OUTPUT_LOG_NAME"
        }
        Write-Log -Message "Finished Sending Email" -Severity Debug
    } 
    catch {
        Write-Log -Message "Could not send email" -Severity Error
    }
}
