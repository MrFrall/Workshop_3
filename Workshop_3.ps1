# 1: Läs in JSON-filen med Get-Content och ConvertFrom-Json

$data = Get-Content -Path "ad_export.json" -Raw -Encoding UTF8 | ConvertFrom-Json

############################################################################################
# Skapa rapport

# $thirtyDaysAgo = (Get-Date).AddDays(-30)
# $inactiveUsers = $data.users | Where-Object {
#    $_.enabled -eq $true -and [datetime]$_.lastLogon -lt $thirtyDaysAgo
#}

# 9: Skapa en funktion Get-InactiveAccounts med parameter för antal dagar


function Get-InactiveAccounts {
    param(
        [int]$Days = 30
    )

    $cutoffDate = (Get-Date).AddDays(-$Days)
    $result = @()

    foreach ($user in $data.users) {
        if ($user.enabled -eq $true) {
            try {
                if ([datetime]$user.lastLogon -lt $cutoffDate) {
                    $result += $user
                }
            }
            catch {
                continue
            }
        }
    }

    return $result
}


$inactive = Get-InactiveAccounts -Days 30

# 3: Lista alla användare som inte loggat in på 30+ dagar
# Uppgift nummer 3 blev överkörd av uppgift nummer 9. Jag har dock sparat koden som jag hade tidigare för jag är osäker på om den ska va med eller ej.
# Tanken kanske var att man skulle ha en extra fil med just den funktionen och sedan anropa den därifrån. Förstod kanske inte riktigt. 

# $report += "{0,-15} {1,-25} {2,-20}`n" -f "Kontonamn", "Namn", "Senast inloggad`n"

# foreach ($user in $inactiveUsers) {
#    $report += "{0,-15} {1,-25} {2}`n" -f $user.samAccountName, $user.displayName, $user.lastLogon
# }
$report = ""

$report += "Lista över inaktiva användare >30 dagar:`n"
$report += "------------------------------------------------------------------`n"
$report += "{0,-15} {1,-25} {2,-20} {3,-10}`n" -f "Kontonamn", "Namn", "Senast inloggad", "Inaktiva dagar"

foreach ($user in $inactive) {
    $daysInactive = ((Get-Date) - [datetime]$user.lastLogon).Days
    $report += "{0,-15} {1,-25} {2,-20} {3,-10}`n" -f $user.samAccountName, $user.displayName, $user.lastLogon, $daysInactive
}

# 4: Räkna antal användare per avdelning med enkel loop

$usersByDept = $data.users | Group-Object department

$report += @"

Användare per avdelning:
------------------------------------------------------------------

"@

$report += "{0,-15} {1,-25}`n" -f "Avdelning", "Antal användare`n"
foreach ($dept in $usersByDept) {
    $report += "{0,-15} {1,-25}`n" -f $dept.Name, $dept.Count
}

# 5: Använd Group-Object för att gruppera datorer per site

$activeComputers = $data.computers | Where-Object { $_.enabled -eq $true }
$compPerSite = $activeComputers | Group-Object site

$report += @"

Antal datorer per avdelning (site):
------------------------------------------------------------------

"@

$report += "{0,-25} {1,-10}`n" -f "Site", "Antal`n"

foreach ($group in $compPerSite) {
    $report += "{0,-25} {1,-10}`n" -f $group.Name, $group.Count
}

$report += "`n"

# 7: Beräkna hur många dagars lösenordsålder varje användare har

$report += "{0,-15} {1,-25}" -f "Användare", "Dagar sedan lösenordsbyte`n"

$report += @"
------------------------------------------------------------------`n
"@

foreach ($user in $data.users | Where-Object { $_.enabled -eq $true }) {
    try {
        $passwordAge = (Get-Date) - [datetime]$user.passwordLastSet
        $report += "{0,-25} {1,-10}`n" -f $user.displayName, $([int]$passwordAge.TotalDays)
    }
    catch {
        continue
    }
}


# 8: Lista de 10 datorer som inte checkat in på längst tid (använd Sort-Object)

$report += "`n"
$report += "Datorer med längst tid sedan senaste incheckning`n"
$report += "{0,-25} {1,-20}`n" -f "Datornamn", "Dagar sedan incheckning"
$report += "------------------------------------------------------------`n"

$gamlaInlogg = $activeComputers | Where-Object { $_.lastLogon -ne $null } | Sort-Object lastLogon | Select-Object -First 10


foreach ($computer in $gamlaInlogg) {
    try {
        $senastInlogg = (Get-Date) - [datetime]$computer.lastLogon
        $report += "{0,-25} {1,-20}`n" -f $computer.name, [int]$senastInlogg.TotalDays
    }
    catch {
        continue
    }
}


# 11: Skapa executive summary med varningar


$snartUtdateradeKonton = @()

foreach ($user in $data.users) {
    if ($user.enabled -eq $true -and $null -ne $user.accountExpires) {
        try {
            if ([datetime]$user.accountExpires -lt (Get-Date).AddDays(30)) {
                $snartUtdateradeKonton += $user
            }
        }
        catch {
            Write-Warning "Kunde inte tolka utgångsdatum för användare: $($user.displayName)"
            continue
        }
    }
}

$gamlaDatorer = @()

foreach ($computer in $data.computers) {
    if ($computer.enabled -eq $true) {
        try {
            if ([datetime]$computer.lastLogon -lt (Get-Date).AddDays(-30)) {
                $gamlaDatorer += $computer
            }
        }
        catch {
            Write-Warning "Kunde inte tolka senaste inloggning för dator: $($computer.name)"
            continue
        }
    }
}

$gamlaLosen = @()

foreach ($user in $data.users) {
    if ($user.enabled -eq $true) {
        try {
            if ([datetime]$user.passwordLastSet -lt (Get-Date).AddDays(-90)) {
                $gamlaLosen += $user
            }
        }
        catch {
            Write-Warning "Kunde inte beräkna lösenordsålder för användare: $($user.displayName)"
            continue  
        }
    }
}



$reportExecutive = @"
EXECUTIVE SUMMARY
==================================================================

⚠ Konton som löper ut inom 30 dagar: $($snartUtdaterade.Count)
⚠ Datorer som inte setts på 30+ dagar: $($gamlaDatorer.Count)
⚠ Användare med lösenord äldre än 90 dagar: $($gamlaLosen.Count)
`n
"@

# 2: Visa domännamn och exportdatum (blev flyttad)


$reportIntro = @"
ACTIVE DIRECTORY AUDIT

==================================================================

Domännamn: $($data.domain)
Exportdatum: $($data.export_date)

Total mängd användare: $($data.users.Count)
Inaktiva användare: $($inactive.count)


"@

# Spara rapport

$reportIntro + $reportExecutive + $report | Out-File -FilePath "ad_report.txt"

# 6: Skapa CSV-fil inactive_users.csv med användare som inte loggat in på 30+ dagar (använd Export-Csv)

$inactive | Select-Object samAccountName, displayName, lastLogon | Export-Csv -Path "inactive_users.csv"  -NoTypeInformation -Encoding UTF8

# Dessa uppgifter ligger inne lite varstans... 
# 10: Generera en professionell textrapport med tydlig formatering och sektioner
# 12: Implementera try/catch för robust felhantering vid datumparsing