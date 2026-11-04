
# ============================================================
# HROMADNY IMPORT ZAKU DO ACTIVE DIRECTORY
# Zdroj: C:\Import\new.xlsx
# ============================================================

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# ------------------------------------------------------------
# KONFIGURACE
# ------------------------------------------------------------

$ImportFolder = "C:\Import"
$ExcelFile    = "$ImportFolder\new.xlsx"
$LogFile      = "$ImportFolder\Import-Zaci.log"

$GroupName    = "gZaci"
$LogonScript  = "logon.cmd"
$HomeDrive    = "U:"
$HomeBase     = "\\Server1\hz"

# Povolené složky tříd
$AllowedOUFolders = @("2024A", "2024B")

# ------------------------------------------------------------
# KONTROLA VSTUPŮ
# ------------------------------------------------------------

if (!(Test-Path $ExcelFile)) {
    Write-Host "CHYBA: Soubor $ExcelFile nebyl nalezen!" -ForegroundColor Red
    exit 1
}

try {
    $Group = Get-ADGroup -Identity $GroupName -ErrorAction Stop
}
catch {
    Write-Host "CHYBA: Skupina $GroupName nebyla nalezena!" -ForegroundColor Red
    exit 1
}

# Načtení Excelu bez záhlaví
$Students = Import-Excel -Path $ExcelFile -NoHeader

if (!$Students -or $Students.Count -eq 0) {
    Write-Host "CHYBA: Excel neobsahuje žádné záznamy!" -ForegroundColor Red
    exit 1
}

# ------------------------------------------------------------
# STATISTIKA
# ------------------------------------------------------------

$SuccessCount = 0
$ErrorCount   = 0
$SkipCount    = 0

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "       IMPORT ZAKU DO ACTIVE DIRECTORY"
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Zdroj: $ExcelFile"
Write-Host "Pocet zaznamu: $($Students.Count)"
Write-Host "============================================" -ForegroundColor Cyan

# ------------------------------------------------------------
# IMPORT UŽIVATELŮ
# ------------------------------------------------------------

foreach ($Student in $Students) {

    $Username    = ([string]$Student.P1).Trim()
    $Password    = [string]$Student.P2
    $DisplayName = ([string]$Student.P3).Trim()
    $FirstName   = ([string]$Student.P4).Trim()
    $LastName    = ([string]$Student.P5).Trim()
    $OUFolder    = ([string]$Student.P6).Trim()
    $OUPath      = ([string]$Student.P7).Trim()

    if ([string]::IsNullOrWhiteSpace($Username)) {
        $SkipCount++
        continue
    }

    Write-Host ""
    Write-Host "Zpracovavam: $Username" -ForegroundColor Cyan

    try {

        # Kontrola povinných údajů
        if (
            [string]::IsNullOrWhiteSpace($Password) -or
            [string]::IsNullOrWhiteSpace($DisplayName) -or
            [string]::IsNullOrWhiteSpace($FirstName) -or
            [string]::IsNullOrWhiteSpace($LastName) -or
            [string]::IsNullOrWhiteSpace($OUFolder) -or
            [string]::IsNullOrWhiteSpace($OUPath)
        ) {
            throw "Chybi nektery z povinnych udaju."
        }

        # Kontrola povolené OU složky
        if ($OUFolder -notin $AllowedOUFolders) {
            throw "Nepovolena slozka OU: $OUFolder"
        }

        # Kontrola existence OU
        Get-ADOrganizationalUnit -Identity $OUPath -ErrorAction Stop |
            Out-Null

        # Kontrola duplicity uživatele
        $ExistingUser = Get-ADUser `
            -Filter "SamAccountName -eq '$Username'" `
            -ErrorAction Stop

        if ($ExistingUser) {
            throw "Uzivatel jiz existuje v Active Directory."
        }

        # Domovská složka
        $HomeDirectory = "$HomeBase\$OUFolder\$Username"

        # Převod hesla
        $SecurePassword = ConvertTo-SecureString `
            $Password `
            -AsPlainText `
            -Force

        # ----------------------------------------------------
        # VYTVOŘENÍ UŽIVATELE
        # ----------------------------------------------------

        New-ADUser `
            -Name $DisplayName `
            -SamAccountName $Username `
            -UserPrincipalName "$Username@skola.intranet" `
            -DisplayName $DisplayName `
            -GivenName $FirstName `
            -Surname $LastName `
            -Path $OUPath `
            -AccountPassword $SecurePassword `
            -Enabled $true `
            -PasswordNeverExpires $true `
            -CannotChangePassword $false `
            -ChangePasswordAtLogon $false `
            -ScriptPath $LogonScript `
            -HomeDirectory $HomeDirectory `
            -HomeDrive $HomeDrive `
            -ErrorAction Stop

        # ----------------------------------------------------
        # ČLENSTVÍ VE SKUPINĚ
        # ----------------------------------------------------

        Add-ADGroupMember `
            -Identity $Group `
            -Members $Username `
            -ErrorAction Stop

        # ----------------------------------------------------
        # LOG
        # ----------------------------------------------------

        $Message = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') OK: $Username"

        Add-Content -Path $LogFile -Value $Message

        $SuccessCount++

        Write-Host "UZIVATEL VYTVOREN" -ForegroundColor Green
        Write-Host "  Login: $Username"
        Write-Host "  Jmeno: $DisplayName"
        Write-Host "  OU: $OUFolder"
        Write-Host "  Skupina: $GroupName"
        Write-Host "  Home: $HomeDirectory"
        Write-Host "  Disk: $HomeDrive"
        Write-Host "  Logon: $LogonScript"

    }
    catch {

        $ErrorMessage = $_.Exception.Message

        $Message = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') CHYBA: $Username - $ErrorMessage"

        Add-Content -Path $LogFile -Value $Message

        $ErrorCount++

        Write-Host "CHYBA: $ErrorMessage" -ForegroundColor Red
    }
}

# ------------------------------------------------------------
# ZÁVĚREČNÉ VYHODNOCENÍ
# ------------------------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "             IMPORT DOKONCEN"
Write-Host "============================================" -ForegroundColor Cyan

Write-Host "Uspesne vytvoreno: $SuccessCount" -ForegroundColor Green
Write-Host "Chyby:             $ErrorCount" -ForegroundColor Red
Write-Host "Preskoceno:        $SkipCount" -ForegroundColor Yellow

Write-Host ""
Write-Host "Log: $LogFile"
Write-Host "============================================" -ForegroundColor Cyan
