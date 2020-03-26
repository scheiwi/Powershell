
function Write-Log() { 

    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$Msg,
        [Parameter(Mandatory = $false)][int]$LogLevel = 1,
        [Parameter(Mandatory = $false)][string]$Color,
        [Parameter(Mandatory = $false)][bool]$isError = $false,
        [switch]$NoNewLine
    )

    $LogColorList = @("green", "yellow", "Cyan", "magenta", "white") # Black, Blue, DarkBlue, Cyan, DarkCyan, Gray, DarkGray, Green, DarkGreen, Magenta, DarkMagenta, Red, DarkRed, White, Yellow, DarkYellow
    if ($LogLevel -gt 5) {$LogLevel = 5}
	
    if ($Color -ne "") {
        $Outcolor = $Color
    }
    else {
        $Outcolor = $LogColorList[$LogLevel - 1]
    }
	
    if ($isError) {
        $Outcolor = "red"
        $Msg = "ERROR: " + $Msg
    }

    $date2 = Get-Date -format dd.MM.yyyy 
    $time = Get-Date -format HH:mm:ss
	
    $NoNewLineSwitch = ""
    if ($NoNewLine) {
        $NoNewLineSwitch = @{NoNewLine = $true}
    }
    
    if ($Global:NextNoNewLine) {
        $Out = ($Msg)
        $Global:NextNoNewLine = $false
    }
    else {
        $Out = ($date2 + " " + $time + "  " * $LogLevel + $Msg)   
    }

    Write-Host $Out -ForegroundColor $Outcolor @NoNewLineSwitch

    [System.IO.File]::AppendAllText($LogFile, $Out)
    if (!$NoNewLine) {
        [System.IO.File]::AppendAllText($LogFile, [System.Environment]::NewLine)
    }

    if ($NoNewLine) {
        $Global:NextNoNewLine = $true
    }
}

function CleanUpLogFiles() {
    # Remvoe all log files which have been created before the calculated clean up threshold
    # Notice that only files with extension "LOG" will be deleted
    $logFilesToDelete = Get-ChildItem -Path ($executingScriptDirectory + "\log") -Filter "*.log" | Where-Object {$_.CreationTime.CompareTo($cleanUptRetentionTime) -eq -1}
		
    Write-log ("Cleaning up log files which are older than {0}" -f $cleanUptRetentionTime)
		
    if ($logFilesToDelete) {
        $logFilesToDelete | Remove-Item
        Write-Log ("{0} log files have been deleted" -f $logFilesToDelete.Count) -LogLevel 2
    }
    else {
        Write-Log "No log files to delete" -LogLevel 2
    }
    Write-Log ""
}

function ReadInConfiguration() {
    if ($ConfigurationFile) {
        #Write-Log ("Loading manually specific Configuration File: " + $ConfigurationFile)
        [xml]$script:xml = Get-Content $ConfigurationFile -Encoding "utf8"
    }
    elseif (Test-Path (Join-Path $executingScriptDirectory ("Configuration-" + $env:computername + ".xml"))) {
        #Write-Log ("Loading Machine specific Configuration File: " + "Configuration-" + $env:computername + ".xml")
        [xml]$script:xml = Get-Content (Join-Path $executingScriptDirectory ("Configuration-" + $env:computername + ".xml")) -Encoding "utf8"
    }
    elseif (Test-Path (Join-Path $executingScriptDirectory "Configuration.xml")) {
        #Write-Log ("Loading general Configuration File: " + "Configuration.xml")
        [xml]$script:xml = Get-Content (Join-Path $executingScriptDirectory "Configuration.xml") -Encoding "utf8"
    }
    elseif (Test-Path (Join-Path $executingScriptDirectory "Configuration-Sample.xml")) {
        #Write-Log ("No Configuration File found, generating Configuration.xml from sample configuration! ")
        Copy-Item -Path (Join-Path $executingScriptDirectory "Configuration-Sample.xml") -Destination (Join-Path $executingScriptDirectory ("Configuration.xml") )
        #Write-Log ("Configuration File: " + "Configuration.xml" + " created, please make sure to adjust the configuration to your needs and restart the script!")
        break
    }
    else {
        throw("No configuration file found!")
    }
	
    <#     #Read in logging configuration
    if ($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Years) {
        $logCleanUpRetentionYears = (-1) * [int]::Parse($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Years)
    }
    else {
        $logCleanUpRetentionYears = 0
    }
    if ($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Months) {
        $logCleanUpRetentionMonths = (-1) * [int]::Parse($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Months)
    }
    else {
        $logCleanUpRetentionMonths = 0
    }
    if ($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Days) {
        $logCleanUpRetentionDays = (-1) * [int]::Parse($xml.Configuration.General.Log.CleanUp.RetentionPeriod.Days)
    }
    else {
        $logCleanUpRetentionDays = 30
    }
	
    $dtNow = (Get-Date).Date
    $script:cleanUptRetentionTime = $dtNow.AddYears($logCleanUpRetentionYears)
    $script:cleanUptRetentionTime = $dtNow.AddMonths($logCleanUpRetentionMonths)
    $script:cleanUptRetentionTime = $dtNow.AddDays($logCleanUpRetentionDays)
	
    Write-Log "" #>
}

function End-Script() {
    Write-Log ""
		
    $rundateend = Get-Date
		
    Write-Log "#################################################################################"
    Write-Log "# End of Script "
    Write-Log ("# Starttime: " + $rundate)
    Write-Log ("# EndTimetime: " + $rundateend)
    Write-Log ("# Run Time: " + ($rundateend - $rundate))
    Write-Log ("# Log File: " + $LogFile)
    Write-Log "#################################################################################"
}

function Load-PSSnapins() {
    Write-Log "Loading PSSnapins..."

    $snapins = Get-ConfigurationValue "General.SnapIns.SnapIn" -AllowNull
	
    if ($snapins) {
        foreach ($pssnapin in $snapins) {
            Write-Log ("Loading Snapin: " + $pssnapin) -LogLevel 2

            if ((Get-PSSnapin | Where-Object { $_.Name -eq $pssnapin }) -eq $null) {
                Add-PSSnapin $pssnapin -ErrorAction SilentlyContinue
                Write-Log "...Done" -LogLevel 3
            }
            else {
                Write-Log ("PSSnapin " + $pssnapin + " already loaded.") -LogLevel 3
            }
        }
    }
    else {
        Write-Log "No PSSnapins defined in Configuration" -LogLevel 2
    }
    Write-Log ""
}

function Import-Modules() {
    Write-Log "Importing Modules..."

    $modules = Get-ConfigurationValue "General.Modules.Module" -AllowNull
	
    if ($modules) {
        foreach ($module in $modules) {
            Write-Log ("Importing Module: " + $module) -LogLevel 2

            if ((Get-Module | Where-Object { $_.Name -eq $module }) -eq $null) {
                if ((get-module -ListAvailable | Where-Object { $_.Name -eq $module }) -eq $null) {
                    throw ("Could not load Module '" + $module + "'")
                }
                else {
                    Import-Module $module
                    Write-Log "...Done" -LogLevel 3
                }
            }
            else {
                Write-Log ("Module " + $module + " already imported.") -LogLevel 3
            }
        }
    }
    else {
        Write-Log "No Modules defined in Configuration" -LogLevel 2
    }
    Write-Log ""
}

function Start-Script() {
    # Set PW Colors
    $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'black')
    $Host.UI.RawUI.ForegroundColor = 'White'
   
    
	
    if ($host.name -ne "Visual Studio Code Host") {
        $Host.PrivateData.ErrorBackgroundColor = $bckgrnd
        $Host.PrivateData.WarningBackgroundColor = $bckgrnd
        $Host.PrivateData.DebugBackgroundColor = $bckgrnd
        $Host.PrivateData.VerboseBackgroundColor = $bckgrnd
    }
    #Only in ISE
    if (($host.name -ne "ConsoleHost") -and ($host.name -ne "Visual Studio Code Host")) {
        $Host.PrivateData.ConsolePaneBackgroundColor = $bckgrnd
    }

    Clear-Host
    ReadInConfiguration

    # Get Script Path
    $global:executingScriptDirectory = $MyInvocation.PSScriptRoot
    #$global:executingScript = $MyInvocation.ScriptName

    # Get Run Date
    $global:rundate = get-date
    $global:rundateformat = get-date -date $rundate -uformat "%Y-%m-%d_%H-%M-%S"
		
    # Set Log File
    $logFolder = $xml.Configuration.General.Log.Folder
    $logPrefix = $xml.Configuration.General.Log.Prefix
    $global:LogFolder = $logFolder
    $global:LogFile = $logFolder + "\" + $logPrefix + "_" + $rundateformat + ".log"
		
    Write-Log "#################################################################################"
    Write-Log "# Starting Script "
    Write-Log ("# Starttime: " + $rundate)
    Write-Log ("# Execution Directory: " + $executingScriptDirectory)
    Write-Log ("# Log File: " + $LogFile)
    Write-Log "#################################################################################"
    Write-Log ""
	
    
    #CleanUpLogFiles
    Load-PSSnapins
    Import-Modules
	
    Write-Log "#################################################################################"
    Write-Log ""

    if ($CreateScheduledTask) {
        Create-ScheduledTask
        break
    }

}

function Get-ConfigurationValue() { 
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$AllowNull
    )

    try {
        $result = Invoke-Expression ("`$xml.Configuration." + $Path)
    }
    catch {
        throw ("XML Node '" + $path + "' does not exist in Configuration")
    }


    if (($result -eq "" -or $result -eq $null) -and $AllowNull -eq $false) {
        throw ("XML Node '" + $path + "' is Null or Empty")
    }
    
    return $result
}

function ExecuteSqlQuery ($Server, $Database, [System.Management.Automation.PSCredential]$Credential, $SQLQuery) {

    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "Server={0};Database={1};User ID={2};Password=`"{3}`";Trusted_Connection=False;" -f $Server, $Database, $Credential.UserName, $Credential.GetNetworkCredential().Password
    try {
        $Connection.Open()
    }
    catch {
        throw "Error opening SQL Connection"
    }

    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $SQLQuery

    $DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $Command
    $Dataset = new-object System.Data.Dataset
    $DataAdapter.Fill($Dataset) | Out-Null
    $Connection.Close()

    return $Dataset.Tables[0]
}


function ExecuteNonQuery ($Server, $Database, [System.Management.Automation.PSCredential]$Credential, $SQLQuery) {

    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "Server={0};Database={1};User ID={2};Password=`"{3}`";Trusted_Connection=False;" -f $Server, $Database, $Credential.UserName, $Credential.GetNetworkCredential().Password
    try {
        $Connection.Open()
    }
    catch {
        throw "Error opening SQL Connection"
    }

    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $SQLQuery
    $Rows = $Command.ExecuteNonQuery()
   
    return $Rows
}

function Send-NotifyMail ($Subject, $Body) {
    $To = @()
    foreach ($recipient in (Get-ConfigurationValue "General.Mail.Recipients")) {
        $To += $recipient.recipient
    }
    $From = Get-ConfigurationValue "General.Mail.FromAdress"
    $SMTPServer = Get-ConfigurationValue "General.Mail.SMTPServer"

    $mailparam = @{
        To         = $To
        From       = $From
        Subject    = $Subject
        Body       = $Body
        SmtpServer = $SMTPServer
    }

    Send-MailMessage @mailparam
}


function Get-SavedCredentials {
    param($credentialName, $Message)
    if (Test-Path (".\Credentials-$credentialName.xml")) {
        Write-Log "Found saved Credentials..." -LogLevel 2
        $Cred = Import-Clixml (".\cred.xml")
		
    }
    else {
        Write-Log "No Credential File found, please insert Credentials to be saved..." -LogLevel 2
        $Cred = Get-Credential -Message $Message
        $Cred | Export-Clixml (".\cred.xml")
    }
    return $Cred
}

function lock-Sitecollection {
    param(
        $Url,
        $Credential
    )
    Write-Log "Connecting to SharePoint online..." -LogLevel 3
    Connect-PnPOnline -Url "https://sbb-admin.sharepoint.com" -Credential $Credential

    Write-Log "locking SiteCollection $url"
    $site = Set-PnPTenantSite -Url $url -LockState NoAccess -Wait
    return


}
function unlock-Sitecollection {
    param(
        $Url,
        $Credential

    )
    Write-Log "Connecting to SharePoint online..." -LogLevel 3

    Connect-PnPOnline -Url "https://sbb-admin.sharepoint.com" -Credential $Credential

    $site = Get-PnPTenantSite -Url $url
    if ($site.LockState -ne "Unlock") {
        Write-Log "unlocking SiteCollection $url"
        Set-PnPTenantSite -Url $url -LockState Unlock -Wait
    }
    else {   
        Write-Log "SiteCollection $url is already unlocked"
    }
    return


}

