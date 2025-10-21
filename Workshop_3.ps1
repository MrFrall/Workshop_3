# 1: Läs in JSON-filen med Get-Content och ConvertFrom-Json

$data = Get-Content -Path "ad_export.json" -Raw -Encoding UTF8 | ConvertFrom-Json


# 2: Visa domännamn och exportdatum

# Domännamn: $($data.domain)
# Exportdatum: $($data.export_date) 


# 3: Lista alla användare som inte loggat in på 30+ dagar
$thirtyDaysAgo = (Get-Date).AddDays(-30)
$inactiveUsers = $data.users | Where-Object {
    $_.enabled -eq $true -and [datetime]$_.lastLogon -lt $thirtyDaysAgo
}


# 4: Räkna antal användare per avdelning med enkel loop

$usersByDept = $data.users | Group-Object department

# 5: Använd Group-Object för att gruppera datorer per site

$activeComputers = $data.computers | Where-Object { $_.enabled -eq $true }
$compPerSite = $activeComputers | Group-Object site

# 6: Skapa CSV-fil inactive_users.csv med användare som inte loggat in på 30+ dagar (använd Export-Csv)


############################################################################################
# Skapa rapport

$report = @"

ACTIVE DIRECTORY AUDIT

==================================================================

Domännamn: $($data.domain)
Exportdatum: $($data.export_date)

Total Users: $($data.users.Count)
Inactive Users $($inactiveUsers.count)

Lista över inaktiva användare >30 dagar:
------------------------------------------------------------------

"@

$report += "{0,-15} {1,-25} {2,-20}`n" -f "Kontonamn", "Namn", "Senast inloggad`n"

foreach ($user in $inactiveUsers) {
    $report += "{0,-15} {1,-25} {2}`n" -f $user.samAccountName, $user.displayName, $user.lastLogon
}

$report += @"

Användare per avdelning:
------------------------------------------------------------------

"@

$report += "{0,-15} {1,-25}`n" -f "Avdelning", "Antal användare`n"
foreach ($dept in $usersByDept) {
    $report += "{0,-15} {1,-25}`n" -f $dept.Name, $dept.Count
}

$report += @"

Antal datorer per avdelning (site):
------------------------------------------------------------------

"@

$report += "{0,-25} {1,-10}`n" -f "Site", "Antal`n"

foreach ($group in $compPerSite) {
    $report += "{0,-25} {1,-10}`n" -f $group.Name, $group.Count
}

$report += "`n"

$report += "{0,-15} {1,-25}" -f "Användare", "Dagar sedan lösenordsbyte`n"

$report += @"
------------------------------------------------------------------`n
"@

foreach ($user in $data.users | Where-Object { $_.enabled -eq $true }) {
    $passwordAge = (Get-Date) - [datetime]$user.passwordLastSet
    $report += "{0,-25} {1,-10}`n" -f $user.displayName, $([int]$passwordAge.TotalDays) 
}


$report += "`n"
$report += "Datorer med längst tid sedan senaste incheckning`n"
$report += "{0,-25} {1,-20}`n" -f "Datornamn", "Dagar sedan incheckning"
$report += "------------------------------------------------------------`n"

$gamlaInlogg = $activeComputers | Where-Object { $_.lastLogon -ne $null } | Sort-Object lastLogon | Select-Object -First 10

foreach ($computer in $gamlaInlogg) {
    $senastInlogg = (Get-Date) - [datetime]$computer.lastLogon
    $report += "{0,-25} {1,-20}`n" -f $computer.name, [int]$senastInlogg.TotalDays
}



# Spara rapport

$report | Out-File -FilePath "ad_report.txt"

$inactiveUsers | Select-Object samAccountName, displayName, lastLogon | Export-Csv -Path "inactive_users.csv"  -NoTypeInformation -Encoding UTF8