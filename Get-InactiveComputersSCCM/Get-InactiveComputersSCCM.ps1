<#
.SYNOPSIS
  Hämtar inaktiva datorer från en AD-OU och visar hårdvaru-/OS-info från ClientHealth-databasen.

.DESCRIPTION
  - Läser konfiguration från Get-InactiveComputersSCCM.xml
  - Löser OU via GUID (undviker teckenkodningsproblem med å/ä/ö)
  - Frågar ClientHealth SQL-databasen med dbatools (Invoke-DbaQuery)
  - Valfritt: exporterar till CSV-fil (-ExportCSV)
  - Valfritt: genererar HTML-rapport med PSWriteHTML (-HTMLReport)
  - Valfritt: skickar HTML-rapport + CSV som bilagor med summary i body via Send-MailKitMessage (-SendMail)

.PARAMETER HTMLReport
  Genererar HTML-rapport och öppnar den i webbläsaren.

.PARAMETER SendMail
  Skickar HTML-rapport + CSV som bilagor med summary i e-postbody.

.PARAMETER DryRun
  Åsidosätter DryRun-inställningen i XML. Genererar HTML men skickar ingen e-post.

.PARAMETER ExportCSV
  Exporterar resultatet till en CSV-fil i HTMLfilePath-mappen.
#>

[CmdletBinding()]
param(
    [switch]$HTMLReport,
    [switch]$SendMail,
    [switch]$DryRun,
    [switch]$ExportCSV
)

$scriptversion = '1.7'
$scriptname    = $MyInvocation.MyCommand.Name

Write-Host "Script: $scriptname  Version: $scriptversion"

# ------------------------------------------------------------------------------------
# Loggfunktion
# ------------------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$LogString,
        [Parameter()][ValidateSet('INFO','WARNING','ERROR','SUCCESS')][string]$Severity = 'INFO'
    )
    Write-Host "[$Severity] $LogString"
}

function Write-LogFile {
    param(
        [Parameter(Mandatory)][string]$LogString,
        [Parameter()][ValidateSet('INFO','WARNING','ERROR','SUCCESS')][string]$Severity = 'INFO'
    )
    try {
        $stamp = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        Add-Content -Path $Logfile -Value "$stamp [$Severity] $LogString"
    } catch {
        Write-Warning "Failed to write to log file '$Logfile': $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------------------------
# Läs konfiguration från XML
# ------------------------------------------------------------------------------------
$xmlPath = Join-Path $PSScriptRoot 'Get-InactiveComputersSCCM.xml'

if (-not (Test-Path -LiteralPath $xmlPath)) {
    Write-Log -LogString "Konfigurationsfil saknas: $xmlPath" -Severity ERROR
    exit 1
}

try {
    [System.Xml.XmlDocument]$xml = Get-Content -Path $xmlPath -Raw -Encoding UTF8 -ErrorAction Stop
    Write-Log -LogString "Konfiguration laddad från $xmlPath" -Severity INFO
}
catch {
    Write-Log -LogString "Kunde inte läsa konfigurationsfil. Fel: $_" -Severity ERROR
    exit 1
}

# Hämta värden från XML
$ouGuid        = $xml.Configuration.ActiveDirectory.OUGUID
$sqlServer     = $xml.Configuration.Database.SQLServer
$database      = $xml.Configuration.Database.DatabaseName
$mailSMTP      = $xml.Configuration.Mail.SMTP
$mailPort      = [int]$xml.Configuration.Mail.Port
$mailFrom      = $xml.Configuration.Mail.From
$subjectPrefix = $xml.Configuration.Mail.SubjectPrefix
$mailCustomer  = $xml.Configuration.Report.CustomerName
$htmlPath      = $xml.Configuration.Report.HTMLfilePath
$LogFile       = $xml.Configuration.Report.LogFile

# DryRun: XML-värde kan åsidosättas av parameter
$dryRunXml = $xml.Configuration.General.DryRun
$isDryRun  = ($DryRun) -or ($dryRunXml -eq 'True')

# Hämta mottagarlista från XML
$recipients = @()
foreach ($r in $xml.Configuration.Mail.Recipients.Recipient) {
    if (-not [string]::IsNullOrWhiteSpace($r.email)) {
        $recipients += $r.email
    }
}

# Visa aktiva lägen
if ($isDryRun)                     { Write-Log -LogString "Mode: DRY RUN (ingen e-post skickas)" -Severity INFO }
if ($HTMLReport)                   { Write-Log -LogString "Mode: HTML REPORT (HTML-fil genereras och öppnas)" -Severity INFO }
if ($ExportCSV)                    { Write-Log -LogString "Mode: EXPORT CSV (CSV-fil genereras)" -Severity INFO }
if ($SendMail -and -not $isDryRun) { Write-Log -LogString "Mode: SEND MAIL (skickas till: $($recipients -join ', '))" -Severity INFO }

# ------------------------------------------------------------------------------------
# Modulhantering
# ------------------------------------------------------------------------------------
function Import-RequiredModule {
    param(
        [Parameter(Mandatory)][string]$ModuleName
    )
    try {
        if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
            if (-not (Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue)) {
                Write-Log -LogString "Modulen '$ModuleName' saknas. Installera med: Install-Module $ModuleName" -Severity ERROR
                Write-LogFile -LogString "Modulen '$ModuleName' saknas. Installera med: Install-Module $ModuleName" -Severity ERROR
                return $false
            }
            Write-Log -LogString "Importerar modul '$ModuleName'" -Severity INFO
            Write-LogFile -LogString "Importerar modul '$ModuleName'" -Severity INFO
            Import-Module $ModuleName -ErrorAction Stop
            Write-Log -LogString "Modul '$ModuleName' importerad" -Severity INFO
            Write-LogFile -LogString "Modul '$ModuleName' importerad" -Severity INFO
        }
        else {
            Write-Log -LogString "Modul '$ModuleName' redan laddad" -Severity INFO
            Write-LogFile -LogString "Modul '$ModuleName' redan laddad" -Severity INFO
        }
        return $true
    }
    catch {
        Write-Log -LogString "Misslyckades att importera '$ModuleName'. Fel: $_" -Severity ERROR
        Write-LogFile -LogString "Misslyckades att importera '$ModuleName'. Fel: $_" -Severity ERROR
        return $false
    }
}

$requiredModules = @('ActiveDirectory', 'dbatools')
if ($HTMLReport -or $SendMail)     { $requiredModules += 'PSWriteHTML' }
if ($SendMail -and -not $isDryRun) { $requiredModules += 'Send-MailKitMessage' }

$allModulesLoaded = $true
foreach ($module in $requiredModules) {
    $loaded           = Import-RequiredModule -ModuleName $module
    $allModulesLoaded = $allModulesLoaded -and $loaded
}

if (-not $allModulesLoaded) {
    Write-Log -LogString "En eller flera moduler kunde inte laddas. Avslutar." -Severity ERROR
    Write-LogFile -LogString "En eller flera moduler kunde inte laddas. Avslutar." -Severity ERROR
    exit 1
}

# ------------------------------------------------------------------------------------
# Steg 1: Lös upp OU via GUID
# ------------------------------------------------------------------------------------
Write-Log -LogString "Löser upp OU via GUID: $ouGuid" -Severity INFO
Write-LogFile -LogString "Löser upp OU via GUID: $ouGuid" -Severity INFO

try {
    $searchBase = (Get-ADOrganizationalUnit -Identity $ouGuid -ErrorAction Stop).DistinguishedName
    Write-Log -LogString "OU: $searchBase" -Severity INFO
    Write-LogFile -LogString "OU: $searchBase" -Severity INFO
}
catch {
    Write-Log -LogString "Kunde inte lösa OU med GUID '$ouGuid'. Fel: $_" -Severity ERROR
    Write-LogFile -LogString "Kunde inte lösa OU med GUID '$ouGuid'. Fel: $_" -Severity ERROR
    exit 1
}

# ------------------------------------------------------------------------------------
# Steg 2: Hämta datorer från AD-OU (rekursivt)
# ------------------------------------------------------------------------------------
Write-Log -LogString "Hämtar datorer från AD..." -Severity INFO
Write-LogFile -LogString "Hämtar datorer från AD..." -Severity INFO

try {
    $inactiveComputers = Get-ADComputer -Filter * `
        -SearchBase $searchBase `
        -SearchScope Subtree `
        -Properties LastLogonDate, DistinguishedName `
        -ErrorAction Stop |
        Select-Object Name, LastLogonDate, DistinguishedName

    Write-Log -LogString "Hittade $($inactiveComputers.Count) datorer i OU." -Severity INFO
    Write-LogFile -LogString "Hittade $($inactiveComputers.Count) datorer i OU." -Severity INFO
}
catch {
    Write-Log -LogString "Kunde inte hämta datorer från AD. Fel: $_" -Severity ERROR
    Write-LogFile -LogString "Kunde inte hämta datorer från AD. Fel: $_" -Severity ERROR
    exit 1
}

if ($inactiveComputers.Count -eq 0) {
    Write-Log -LogString "Inga datorer hittades i OU. Avslutar." -Severity WARNING
    Write-LogFile -LogString "Inga datorer hittades i OU. Avslutar." -Severity WARNING
    exit 0
}

# ------------------------------------------------------------------------------------
# Steg 3: Bygg SQL-fråga med IN (...) för alla datorer i ett anrop
# ------------------------------------------------------------------------------------
Set-DbatoolsConfig -FullName sql.connection.trustcert -Value $true

$computerList = ($inactiveComputers.Name | ForEach-Object { "'" + ($_.Replace("'","''")) + "'" }) -join ", "

$query = @"
SELECT
    [Hostname],
    [OperatingSystem],
    [Model],
    [InstallDate],
    [LastLoggedOnUser],
    [LastBootTime]
FROM
    [dbo].[Clients]
WHERE
    [Hostname] IN ($computerList)
"@

# ------------------------------------------------------------------------------------
# Steg 4: Fråga ClientHealth-databasen
# ------------------------------------------------------------------------------------
Write-Log -LogString "Frågar ClientHealth-databasen på $sqlServer ($database)..." -Severity INFO
Write-LogFile -LogString "Frågar ClientHealth-databasen på $sqlServer ($database)..." -Severity INFO

try {
    $clientHealthData = Invoke-DbaQuery -SqlInstance $sqlServer -Database $database -Query $query -ErrorAction Stop
    Write-Log -LogString "SQL-fråga lyckades. $($clientHealthData.Count) rader returnerade." -Severity INFO
    Write-LogFile -LogString "SQL-fråga lyckades. $($clientHealthData.Count) rader returnerade." -Severity INFO
}
catch {
    Write-Log -LogString "SQL-fråga misslyckades. Fel: $_" -Severity ERROR
    Write-LogFile -LogString "SQL-fråga misslyckades. Fel: $_" -Severity ERROR
    exit 1
}

# ------------------------------------------------------------------------------------
# Steg 5: Slå ihop AD och ClientHealth-data + beräkna DaysInactive
# ------------------------------------------------------------------------------------
$clientHealthLookup = @{}
foreach ($item in $clientHealthData) {
    if (-not [string]::IsNullOrWhiteSpace($item.Hostname)) {
        $clientHealthLookup[$item.Hostname] = $item
    }
}

$results = foreach ($computer in $inactiveComputers) {
    $ch = $null
    $null = $clientHealthLookup.TryGetValue($computer.Name, [ref]$ch)

    $daysInactive = if ($computer.LastLogonDate) {
        [math]::Round(((Get-Date) - $computer.LastLogonDate).TotalDays, 0)
    }
    else {
        $null
    }

    [PSCustomObject]@{
        Hostname         = $computer.Name
        LastLogonAD      = $computer.LastLogonDate
        DaysInactive     = $daysInactive
        Model            = if ($ch) { $ch.Model }            else { "Ej i ClientHealth" }
        OperatingSystem  = if ($ch) { $ch.OperatingSystem }  else { "Ej i ClientHealth" }
        InstallDate      = if ($ch) { $ch.InstallDate }      else { $null }
        LastLoggedOnUser = if ($ch) { $ch.LastLoggedOnUser } else { $null }
        LastBootTime     = if ($ch) { $ch.LastBootTime }     else { $null }
    }
}

Write-Log -LogString "Datasammanslagning klar. $($results.Count) poster." -Severity INFO
Write-LogFile -LogString "Datasammanslagning klar. $($results.Count) poster." -Severity INFO

# ------------------------------------------------------------------------------------
# Steg 6: Beräkna summary-data
# ------------------------------------------------------------------------------------
$totalComputers = $results.Count
$notInCH        = ($results | Where-Object { $_.Model -eq "Ej i ClientHealth" }).Count
$inCH           = $totalComputers - $notInCH

$modelSummary = $results |
    Where-Object { $_.Model -ne "Ej i ClientHealth" } |
    Group-Object -Property Model |
    Sort-Object -Property Count -Descending |
    Select-Object @{N='Model';E={$_.Name}}, @{N='Antal';E={$_.Count}}

# Inaktivitetssammanfattning
$withDays = $results | Where-Object { $null -ne $_.DaysInactive }

$inactive_0_30    = ($withDays | Where-Object { $_.DaysInactive -ge 0   -and $_.DaysInactive -le 30  }).Count
$inactive_31_60   = ($withDays | Where-Object { $_.DaysInactive -ge 31  -and $_.DaysInactive -le 60  }).Count
$inactive_61_90   = ($withDays | Where-Object { $_.DaysInactive -ge 61  -and $_.DaysInactive -le 90  }).Count
$inactive_91_180  = ($withDays | Where-Object { $_.DaysInactive -ge 91  -and $_.DaysInactive -le 180 }).Count
$inactive_181_365 = ($withDays | Where-Object { $_.DaysInactive -ge 181 -and $_.DaysInactive -le 365 }).Count
$inactive_365plus = ($withDays | Where-Object { $_.DaysInactive -gt 365 }).Count
$inactiveUnknown  = ($results  | Where-Object { $null -eq $_.DaysInactive }).Count

# ------------------------------------------------------------------------------------
# Steg 8: Exportera till CSV (-ExportCSV eller -SendMail)
# ------------------------------------------------------------------------------------
$csvFile = $null

if ($ExportCSV -or $SendMail) {
    if (-not (Test-Path -LiteralPath $htmlPath)) {
        New-Item -ItemType Directory -Path $htmlPath -Force | Out-Null
    }

    $csvFile = Join-Path $htmlPath ("InaktivaDatorer_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))

    try {
        $results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-Log -LogString "CSV exporterad: $csvFile" -Severity SUCCESS
        Write-LogFile -LogString "CSV exporterad: $csvFile" -Severity SUCCESS
    }
    catch {
        Write-Log -LogString "Misslyckades att exportera CSV. Fel: $_" -Severity ERROR
        Write-LogFile -LogString "Misslyckades att exportera CSV. Fel: $_" -Severity ERROR
        $csvFile = $null
    }
}

# ------------------------------------------------------------------------------------
# Steg 9: HTML-rapport (valfritt)
# ------------------------------------------------------------------------------------
$htmlFile    = $null
$htmlContent = $null
$now         = Get-Date -Format "yyyy-MM-dd HH:mm"
$reportTitle = "Inaktiva datorer - ClientHealth-rapport"

if ($HTMLReport -or $SendMail) {
    $htmlContent = New-HTML -TitleText $reportTitle {

        New-HTMLTag -Tag 'style' {
@"
table { border-collapse: collapse; }
table, th, td { border: 1px solid #cccccc; }
th, td { padding: 4px 8px; font-size: 11px; }
.header-block { margin: 10px auto; font-size: 12px; max-width: 1100px; }
.header-block div { margin: 2px 0; }
.header-label { font-weight: bold; }
"@
        }

        New-HTMLSection -HeaderTextAlignment center -HeaderTextSize 20 -HeaderBackGroundColor DarkBlue -HeaderText $reportTitle {
            New-HTMLTag -Tag 'div' -Attributes @{ class = 'header-block' } {
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Rapport genererad:</span> $now" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Kund:</span> $mailCustomer" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>OU:</span> $searchBase" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Totalt inaktiva datorer:</span> <b>$totalComputers</b>" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Hittade i ClientHealth:</span> <b>$inCH</b>" }
                New-HTMLTag -Tag 'div' { "<span class='header-label'>Ej i ClientHealth:</span> <b>$notInCH</b>" }
            }
        }

        New-HTMLSection -HeaderBackGroundColor DarkBlue -HeaderText "Inaktivitet (dagar)" {
            $inactivityTable = @(
                [PSCustomObject]@{ Intervall = '0-30 dagar';    Antal = $inactive_0_30 }
                [PSCustomObject]@{ Intervall = '31-60 dagar';   Antal = $inactive_31_60 }
                [PSCustomObject]@{ Intervall = '61-90 dagar';   Antal = $inactive_61_90 }
                [PSCustomObject]@{ Intervall = '91-180 dagar';  Antal = $inactive_91_180 }
                [PSCustomObject]@{ Intervall = '181-365 dagar'; Antal = $inactive_181_365 }
                [PSCustomObject]@{ Intervall = '365+ dagar';    Antal = $inactive_365plus }
                [PSCustomObject]@{ Intervall = 'Okänd';         Antal = $inactiveUnknown }
            )
            New-HTMLTable -DataTable $inactivityTable -PagingLength 20
        }

        New-HTMLSection -HeaderBackGroundColor DarkBlue -HeaderText "Datorer" {
            if ($results.Count -eq 0) {
                New-HTMLText -Text "Inga datorer hittades." -Color Red -FontSize 14
            }
            else {
                New-HTMLTable -PagingLength 100 -DataTable $results -ScrollX {
                    New-TableCondition -Name 'Model' -ComparisonType string -Operator eq `
                        -Value 'Ej i ClientHealth' -BackgroundColor '#ffe0e0' -Color '#cc0000'
                }
            }
        }
    }

    if (-not (Test-Path -LiteralPath $htmlPath)) {
        New-Item -ItemType Directory -Path $htmlPath -Force | Out-Null
    }

    $htmlFile = Join-Path $htmlPath ("InaktivaDatorer_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
    $htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8
    Write-Log -LogString "HTML-rapport sparad: $htmlFile" -Severity INFO
    Write-LogFile -LogString "HTML-rapport sparad: $htmlFile" -Severity INFO
}

# ------------------------------------------------------------------------------------
# Steg 10: Skicka e-post (valfritt)
# ------------------------------------------------------------------------------------
if ($SendMail) {
    if ($isDryRun) {
        Write-Log -LogString "DryRun aktiv - e-post skickas inte." -Severity WARNING
        Write-LogFile -LogString "DryRun aktiv - e-post skickas inte." -Severity WARNING
    }
    else {
        if (-not $htmlFile -or -not (Test-Path -LiteralPath $htmlFile)) {
            Write-Log -LogString "HTML-fil saknas, kan inte skicka e-post." -Severity ERROR
            Write-LogFile -LogString "HTML-fil saknas, kan inte skicka e-post." -Severity ERROR
            exit 1
        }

        $attachments = @($htmlFile)
        if ($csvFile -and (Test-Path -LiteralPath $csvFile)) {
            $attachments += $csvFile
            Write-Log -LogString "CSV bifogas mailet: $csvFile" -Severity INFO
            Write-LogFile -LogString "CSV bifogas mailet: $csvFile" -Severity INFO
        }

        $modelRows = ($modelSummary | ForEach-Object {
            "<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>$($_.Model)</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$($_.Antal)</td></tr>"
        }) -join ""

        $inactiveRows = @"
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>0-30 dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_0_30</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>31-60 dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_31_60</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>61-90 dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_61_90</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>91-180 dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_91_180</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>181-365 dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_181_365</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>365+ dagar</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactive_365plus</td></tr>
<tr><td style='padding:4px 12px;border:1px solid #cccccc;'>Okänd</td><td style='padding:4px 12px;border:1px solid #cccccc;text-align:center;'>$inactiveUnknown</td></tr>
"@

        $mailBody = @"
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; font-size: 13px; color: #222; }
    h2   { color: #003366; }
    h3   { color: #003366; margin-top: 24px; }
    table { border-collapse: collapse; margin-top: 8px; }
    th   { background-color: #003366; color: #ffffff; padding: 6px 12px; text-align: left; border: 1px solid #cccccc; }
    td   { padding: 4px 12px; border: 1px solid #cccccc; }
    .stat-label { font-weight: bold; width: 280px; }
    .note { font-size: 11px; color: #666; margin-top: 20px; }
  </style>
</head>
<body>
  <h2>$reportTitle</h2>

  <p>Rapport genererad: <b>$now</b><br>
     Kund: <b>$mailCustomer</b><br>
     OU: <b>$searchBase</b></p>

  <h3>Sammanfattning</h3>
  <table>
    <tr><td class='stat-label'>Totalt inaktiva datorer i OU</td><td><b>$totalComputers</b></td></tr>
    <tr><td class='stat-label'>Hittade i ClientHealth</td><td><b>$inCH</b></td></tr>
    <tr><td class='stat-label'>Ej registrerade i ClientHealth</td><td><b>$notInCH</b></td></tr>
  </table>

  <h3>Inaktivitet (dagar)</h3>
  <table>
    <tr><th>Intervall</th><th>Antal</th></tr>
    $inactiveRows
  </table>

  <h3>Modeller (från ClientHealth)</h3>
  <table>
    <tr><th>Modell</th><th>Antal</th></tr>
    $modelRows
  </table>

  <p class='note'>Bilagor: HTML-rapport med fullständig datortabell samt CSV-fil för import i Excel.</p>
</body>
</html>
"@

        $subject     = "$subjectPrefix - $now"
        $fromAddress = New-Object MimeKit.MailboxAddress ('', $mailFrom)

        foreach ($addr in $recipients) {
            $toAddress = New-Object MimeKit.MailboxAddress ('', $addr)

            Write-Log -LogString "Skickar e-post till $addr..." -Severity INFO
            Write-LogFile -LogString "Skickar e-post till $addr..." -Severity INFO

            try {
                Send-MailKitMessage `
                    -SMTPServer     $mailSMTP `
                    -Port           $mailPort `
                    -From           $fromAddress `
                    -RecipientList  $toAddress `
                    -Subject        $subject `
                    -HTMLBody       $mailBody `
                    -AttachmentList $attachments

                Write-Log -LogString "E-post skickad till $addr." -Severity SUCCESS
                Write-LogFile -LogString "E-post skickad till $addr." -Severity SUCCESS
            }
            catch {
                Write-Log -LogString "Misslyckades att skicka e-post till $addr. Fel: $_" -Severity ERROR
                Write-LogFile -LogString "Misslyckades att skicka e-post till $addr. Fel: $_" -Severity ERROR
            }
        }
    }
}

Write-Log -LogString "Klart." -Severity SUCCESS
Write-LogFile -LogString "Klart." -Severity SUCCESS
