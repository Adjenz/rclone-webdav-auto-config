#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Script de deploiement automatique de rclone avec montage WebDAV
.DESCRIPTION
    Ce script installe automatiquement rclone, WinFSP, NSSM et configure un service de montage WebDAV
    Version amelioree avec validation des entrees et gestion d'erreurs renforcee
.NOTES
    Auteur: Script de deploiement rclone
    Version: 1.5
    Necessite les droits administrateur et winget
.PARAMETER DryRun
    Execute le script en mode simulation sans effectuer de changements
#>

param(
    [switch]$DryRun
)

# Configuration des couleurs pour l'affichage
$Host.UI.RawUI.WindowTitle = "Deploiement automatique rclone"

# Variables globales
$script:LogFile = "C:\rclone\deploy-script.log"
$script:ErrorCount = 0

function Write-ColorOutput($ForegroundColor, $Message) {
    Write-Host $Message -ForegroundColor $ForegroundColor
    Add-LogEntry $Message
}

function Write-Header($Message) {
    Write-Host "`n" + "="*60 -ForegroundColor Yellow
    Write-Host $Message -ForegroundColor Yellow
    Write-Host "="*60 -ForegroundColor Yellow
    Add-LogEntry ("="*60)
    Add-LogEntry $Message
    Add-LogEntry ("="*60)
}

function Add-LogEntry($Message) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    
    # Creer le dossier si necessaire
    $logDir = Split-Path $script:LogFile -Parent
    if (!(Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    # Ajouter au fichier log
    Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-Winget {
    try {
        $null = Get-Command winget -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Test-NetworkConnection {
    try {
        $null = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        return $true
    } catch {
        return $false
    }
}

function Install-Winget {
    if (!(Test-Winget)) {
        Write-ColorOutput "Yellow" "Winget n'est pas disponible. Installation en cours..."
        
        if ($DryRun) {
            Write-ColorOutput "Cyan" "[DRY RUN] Installation de winget simulee"
            return
        }
        
        try {
            $progressPreference = 'silentlyContinue'
            Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            Add-AppxPackage "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            Remove-Item "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -Force
            Write-ColorOutput "Green" "Winget installe avec succes."
        } catch {
            Write-ColorOutput "Red" "Impossible d'installer winget automatiquement."
            Write-ColorOutput "Yellow" "Veuillez installer manuellement Microsoft App Installer depuis le Microsoft Store."
            throw "Winget requis pour continuer"
        }
    } else {
        Write-ColorOutput "Green" "Winget deja disponible."
    }
}

function Install-Dependencies {
    Write-Header "INSTALLATION DES DEPENDANCES"
    
    # Verifier la connexion reseau
    if (!(Test-NetworkConnection)) {
        throw "Aucune connexion reseau detectee. Verifiez votre connexion Internet."
    }
    
    Write-ColorOutput "Yellow" "Installation de rclone..."
    if ($DryRun) {
        Write-ColorOutput "Cyan" "[DRY RUN] Installation de rclone simulee"
    } else {
        try {
            $result = winget install Rclone.Rclone --accept-package-agreements --accept-source-agreements --silent 2>&1
            Write-ColorOutput "Green" "rclone installe avec succes."
        } catch {
            Write-ColorOutput "Red" "Erreur lors de l'installation de rclone: $($_.Exception.Message)"
            $script:ErrorCount++
        }
    }
    
    Write-ColorOutput "Yellow" "Installation de WinFSP..."
    if ($DryRun) {
        Write-ColorOutput "Cyan" "[DRY RUN] Installation de WinFSP simulee"
    } else {
        try {
            $result = winget install WinFsp.WinFsp --accept-package-agreements --accept-source-agreements --silent 2>&1
            Write-ColorOutput "Green" "WinFSP installe avec succes."
        } catch {
            Write-ColorOutput "Red" "Erreur lors de l'installation de WinFSP: $($_.Exception.Message)"
            $script:ErrorCount++
        }
    }
    
    Write-ColorOutput "Yellow" "Installation de NSSM..."
    $nssmPath = "C:\nssm"
    if (!(Test-Path $nssmPath)) {
        if ($DryRun) {
            Write-ColorOutput "Cyan" "[DRY RUN] Installation de NSSM simulee"
        } else {
            New-Item -Path $nssmPath -ItemType Directory -Force | Out-Null
            $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
            $nssmZip = "$env:TEMP\nssm.zip"
            
            try {
                Write-ColorOutput "Yellow" "Telechargement de NSSM..."
                $progressPreference = 'silentlyContinue'
                
                # Ajouter un timeout et retry
                $maxRetries = 3
                $retryCount = 0
                $downloaded = $false
                
                while (!$downloaded -and $retryCount -lt $maxRetries) {
                    try {
                        Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip -TimeoutSec 30
                        $downloaded = $true
                    } catch {
                        $retryCount++
                        if ($retryCount -lt $maxRetries) {
                            Write-ColorOutput "Yellow" "Echec du telechargement, nouvelle tentative ($retryCount/$maxRetries)..."
                            Start-Sleep -Seconds 2
                        } else {
                            throw
                        }
                    }
                }
                
                Write-ColorOutput "Yellow" "Extraction de NSSM..."
                Expand-Archive -Path $nssmZip -DestinationPath $env:TEMP -Force
                
                $nssmSource = "$env:TEMP\nssm-2.24\win64"
                Copy-Item -Path "$nssmSource\*" -Destination $nssmPath -Force
                
                Remove-Item $nssmZip -Force
                Remove-Item "$env:TEMP\nssm-2.24" -Recurse -Force
                
                Write-ColorOutput "Green" "NSSM installe avec succes."
            } catch {
                Write-ColorOutput "Red" "Erreur lors de l'installation de NSSM: $($_.Exception.Message)"
                $script:ErrorCount++
            }
        }
    } else {
        Write-ColorOutput "Green" "NSSM deja installe."
    }
    
    if (!$DryRun) {
        $currentPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
        if ($currentPath -notlike "*$nssmPath*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$nssmPath", [EnvironmentVariableTarget]::Machine)
            $env:Path += ";$nssmPath"
        }
        
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }
}

function Test-DriveLetter($Letter) {
    $usedLetters = (Get-PSDrive -PSProvider FileSystem).Name
    return !($usedLetters -contains $Letter)
}

function Test-ValidUrl($Url) {
    return $Url -match "^https?://[\w\-\.]+(:\d+)?(/.*)?$"
}

function Test-ValidRemoteName($Name) {
    # Verifier les caracteres valides pour un nom de remote rclone
    return $Name -match "^[a-zA-Z0-9_\-]+$"
}

function Get-UserConfiguration {
    Write-Header "CONFIGURATION DU REMOTE WEBDAV"
    
    $config = @{}
    
    # Nom du remote avec validation
    do {
        $config.RemoteName = Read-Host "Nom du remote rclone (ex: monserveur)"
        if (!(Test-ValidRemoteName $config.RemoteName)) {
            Write-ColorOutput "Red" "Le nom du remote ne doit contenir que des lettres, chiffres, tirets et underscores."
        }
    } while (!(Test-ValidRemoteName $config.RemoteName))
    
    # URL WebDAV avec validation
    do {
        $config.WebDAVUrl = Read-Host "URL du serveur WebDAV (ex: https://monserveur.com/webdav)"
        if (!(Test-ValidUrl $config.WebDAVUrl)) {
            Write-ColorOutput "Red" "L'URL doit commencer par http:// ou https:// et etre valide."
        }
    } while (!(Test-ValidUrl $config.WebDAVUrl))
    
    $config.Username = Read-Host "Nom d'utilisateur WebDAV"
    
    $securePassword = Read-Host "Mot de passe WebDAV" -AsSecureString
    $config.Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
    
    # Lettre de lecteur avec validation
    do {
        $config.DriveLetter = Read-Host "Lettre de lecteur pour le montage (ex: Z)"
        $config.DriveLetter = $config.DriveLetter.ToUpper()
        
        if ($config.DriveLetter -notmatch "^[D-Z]$") {
            Write-ColorOutput "Red" "Veuillez choisir une lettre entre D et Z."
        } elseif (!(Test-DriveLetter $config.DriveLetter)) {
            Write-ColorOutput "Red" "La lettre $($config.DriveLetter) est deja utilisee. Choisissez une autre lettre."
            Write-ColorOutput "Yellow" "Lettres disponibles: $((Get-AvailableDriveLetters) -join ', ')"
        }
    } while ($config.DriveLetter -notmatch "^[D-Z]$" -or !(Test-DriveLetter $config.DriveLetter))
    
    Write-ColorOutput "Cyan" "`nChoisissez le type de serveur WebDAV:"
    Write-ColorOutput "White" "1. Generique (par defaut)"
    Write-ColorOutput "White" "2. Nextcloud"
    Write-ColorOutput "White" "3. Owncloud"
    Write-ColorOutput "White" "4. Sharepoint"
    Write-ColorOutput "White" "5. Autre"
    
    $vendorChoice = Read-Host "Votre choix (1-5)"
    switch ($vendorChoice) {
        "2" { $config.Vendor = "nextcloud" }
        "3" { $config.Vendor = "owncloud" }
        "4" { $config.Vendor = "sharepoint" }
        "5" { $config.Vendor = Read-Host "Specifiez le vendor" }
        default { $config.Vendor = "other" }
    }
    
    return $config
}

function Get-AvailableDriveLetters {
    $allLetters = 68..90 | ForEach-Object { [char]$_ }  # D to Z
    $usedLetters = (Get-PSDrive -PSProvider FileSystem).Name
    return $allLetters | Where-Object { $usedLetters -notcontains $_ }
}

function Create-RcloneConfig($config) {
    Write-Header "CREATION DE LA CONFIGURATION RCLONE"
    
    $rcloneDir = "C:\rclone"
    $configPath = "$rcloneDir\rclone.conf"
    
    if (!(Test-Path $rcloneDir)) {
        if (!$DryRun) {
            New-Item -Path $rcloneDir -ItemType Directory -Force | Out-Null
        }
        Write-ColorOutput "Green" "Dossier $rcloneDir cree."
    }
    
    # Sauvegarder la configuration existante si presente
    if ((Test-Path $configPath) -and !$DryRun) {
        $backupPath = "$configPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Copy-Item -Path $configPath -Destination $backupPath
        Write-ColorOutput "Yellow" "Configuration existante sauvegardee: $backupPath"
    }
    
    Write-ColorOutput "Yellow" "Chiffrement du mot de passe..."
    if ($DryRun) {
        Write-ColorOutput "Cyan" "[DRY RUN] Creation de la configuration simulee"
        $encryptedPassword = "ENCRYPTED_PASSWORD_SIMULATION"
    } else {
        try {
            $rcloneExe = Get-RcloneExecutable
            $encryptedPassword = & $rcloneExe obscure $config.Password
        } catch {
            Write-ColorOutput "Yellow" "Impossible de chiffrer le mot de passe, utilisation en clair"
            $encryptedPassword = $config.Password
        }
    }
    
    $configContent = @"
[$($config.RemoteName)]
type = webdav
url = $($config.WebDAVUrl)
vendor = $($config.Vendor)
user = $($config.Username)
pass = $encryptedPassword
"@
    
    if (!$DryRun) {
        $configContent | Out-File -FilePath $configPath -Encoding UTF8
    }
    
    Write-ColorOutput "Green" "Configuration rclone creee: $configPath"
    
    return $configPath
}

function Get-RcloneExecutable {
    # Recherche dans le PATH en premier
    try {
        $rcloneInPath = Get-Command rclone.exe -ErrorAction Stop
        if ($rcloneInPath) {
            return $rcloneInPath.Source
        }
    } catch {
        # Continue la recherche
    }
    
    $possiblePaths = @(
        "C:\Program Files\Rclone\rclone.exe",
        "C:\Program Files (x86)\Rclone\rclone.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Rclone.Rclone_*\rclone.exe",
        "C:\Users\$env:USERNAME\AppData\Local\Microsoft\WinGet\Packages\Rclone.Rclone_*\rclone.exe",
        "$env:ProgramFiles\Rclone\rclone.exe",
        "$env:ProgramFiles(x86)\Rclone\rclone.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if ($path -like "*\*") {
            $expandedPaths = Get-ChildItem $path -ErrorAction SilentlyContinue
            if ($expandedPaths) {
                return $expandedPaths[0].FullName
            }
        }
    }
    
    throw "Impossible de trouver rclone.exe. Verifiez que rclone est correctement installe."
}

function Create-RcloneService($config, $configPath) {
    Write-Header "CREATION DU SERVICE WINDOWS"
    
    $serviceName = "rclone_$($config.RemoteName)"
    
    if ($DryRun) {
        Write-ColorOutput "Cyan" "[DRY RUN] Creation du service $serviceName simulee"
        return $serviceName
    }
    
    $rcloneExe = Get-RcloneExecutable
    $logPath = "C:\rclone\rclone.log"
    
    # Construire les arguments avec des guillemets appropries
    $arguments = @(
        "mount",
        "$($config.RemoteName):",
        "$($config.DriveLetter):",
        "--no-console",
        "--vfs-cache-mode", "full",
        "--network-mode",
        "--dir-cache-time", "10s",
        "--config", "`"$configPath`"",
        "--log-file", "`"$logPath`"",
        "--log-level", "INFO"
    ) -join " "
    
    # Verifier si le service existe deja
    $existingService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-ColorOutput "Yellow" "Le service $serviceName existe deja."
        $response = Read-Host "Voulez-vous le remplacer? (O/N)"
        if ($response -ne "O" -and $response -ne "o") {
            throw "Installation annulee par l'utilisateur."
        }
        
        Write-ColorOutput "Yellow" "Arret et suppression du service existant..."
        Stop-Service $serviceName -Force -ErrorAction SilentlyContinue
        & C:\nssm\nssm.exe remove $serviceName confirm 2>&1 | Out-Null
    }
    
    Write-ColorOutput "Yellow" "Creation du nouveau service: $serviceName"
    $installResult = & C:\nssm\nssm.exe install $serviceName "`"$rcloneExe`"" $arguments 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "Erreur lors de la creation du service NSSM: $installResult"
    }
    
    # Configuration du service
    & C:\nssm\nssm.exe set $serviceName AppDirectory "C:\rclone" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName DisplayName "Rclone Mount - $($config.RemoteName)" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName Description "Service de montage rclone pour $($config.RemoteName)" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName Start SERVICE_AUTO_START 2>&1 | Out-Null
    
    # Configuration des actions en cas d'echec
    & C:\nssm\nssm.exe set $serviceName AppExit Default Restart 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName AppRestartDelay 5000 2>&1 | Out-Null
    
    Write-ColorOutput "Green" "Service $serviceName cree avec succes."
    
    return $serviceName
}

function Start-RcloneService($serviceName) {
    Write-Header "DEMARRAGE DU SERVICE"
    
    if ($DryRun) {
        Write-ColorOutput "Cyan" "[DRY RUN] Demarrage du service $serviceName simule"
        return
    }
    
    Write-ColorOutput "Yellow" "Demarrage du service $serviceName..."
    try {
        Start-Service $serviceName
        
        # Attendre le demarrage avec timeout
        $timeout = 30
        $elapsed = 0
        
        while ($elapsed -lt $timeout) {
            Start-Sleep -Seconds 2
            $elapsed += 2
            
            $service = Get-Service $serviceName
            if ($service.Status -eq "Running") {
                Write-ColorOutput "Green" "Service demarre avec succes!"
                
                # Verifier que le lecteur est monte
                Start-Sleep -Seconds 3
                if (Test-Path "$($config.DriveLetter):\") {
                    Write-ColorOutput "Green" "Lecteur $($config.DriveLetter): monte avec succes!"
                } else {
                    Write-ColorOutput "Yellow" "Le lecteur n'est pas encore visible. Cela peut prendre quelques secondes."
                }
                
                return
            }
        }
        
        Write-ColorOutput "Red" "Le service n'a pas demarre dans le delai imparti."
        Write-ColorOutput "Yellow" "Verifiez le fichier log: C:\rclone\rclone.log"
        
    } catch {
        Write-ColorOutput "Red" "Erreur lors du demarrage: $($_.Exception.Message)"
        Write-ColorOutput "Yellow" "Verifiez le fichier log: C:\rclone\rclone.log"
        $script:ErrorCount++
    }
}

function Show-Summary($config, $serviceName) {
    Write-Header "RESUME DE L'INSTALLATION"
    
    if ($DryRun) {
        Write-ColorOutput "Cyan" "MODE DRY RUN - Aucune modification effectuee"
    }
    
    Write-ColorOutput "Green" "rclone installe"
    Write-ColorOutput "Green" "WinFSP installe" 
    Write-ColorOutput "Green" "NSSM installe"
    Write-ColorOutput "Green" "Configuration creee"
    Write-ColorOutput "Green" "Service Windows cree: $serviceName"
    
    Write-ColorOutput "Cyan" "`nInformations:"
    Write-ColorOutput "White" "Remote: $($config.RemoteName)"
    Write-ColorOutput "White" "URL: $($config.WebDAVUrl)"
    Write-ColorOutput "White" "Lecteur: $($config.DriveLetter):"
    Write-ColorOutput "White" "Service: $serviceName"
    Write-ColorOutput "White" "Log rclone: C:\rclone\rclone.log"
    Write-ColorOutput "White" "Log script: $script:LogFile"
    Write-ColorOutput "White" "Config: C:\rclone\rclone.conf"
    
    Write-ColorOutput "Yellow" "`nPour gerer le service:"
    Write-ColorOutput "White" "Demarrer: Start-Service $serviceName"
    Write-ColorOutput "White" "Arreter: Stop-Service $serviceName"
    Write-ColorOutput "White" "Status: Get-Service $serviceName"
    Write-ColorOutput "White" "Logs: Get-Content C:\rclone\rclone.log -Tail 50"
    Write-ColorOutput "White" "Supprimer: C:\nssm\nssm.exe remove $serviceName"
    
    if ($script:ErrorCount -gt 0) {
        Write-ColorOutput "Yellow" "`nATTENTION: $($script:ErrorCount) erreur(s) rencontree(s) durant l'installation."
        Write-ColorOutput "Yellow" "Consultez le fichier log pour plus de details: $script:LogFile"
    } else {
        Write-ColorOutput "Green" "`nInstallation terminee avec succes!"
    }
}

# ===== SCRIPT PRINCIPAL =====

try {
    Clear-Host
    Write-ColorOutput "Green" "Script de deploiement automatique rclone WebDAV"
    Write-ColorOutput "Green" "Version 1.5 - Version amelioree avec validation"
    
    if ($DryRun) {
        Write-ColorOutput "Cyan" "MODE DRY RUN ACTIVE - Aucune modification ne sera effectuee"
    }
    
    Add-LogEntry "Debut du script de deploiement rclone v1.5"
    
    if (!(Test-Administrator)) {
        Write-ColorOutput "Red" "Ce script necessite les droits administrateur."
        Write-ColorOutput "Yellow" "Relancez PowerShell en tant qu'administrateur."
        Read-Host "Appuyez sur Entree pour fermer"
        exit 1
    }
    
    Install-Winget
    Install-Dependencies
    
    $config = Get-UserConfiguration
    $configPath = Create-RcloneConfig $config
    $serviceName = Create-RcloneService $config $configPath
    
    Start-RcloneService $serviceName
    Show-Summary $config $serviceName
    
    Add-LogEntry "Script termine avec $($script:ErrorCount) erreur(s)"
    
} catch {
    Write-ColorOutput "Red" "Erreur durant l'installation: $($_.Exception.Message)"
    Write-ColorOutput "Yellow" "Consultez les logs pour plus d'informations: $script:LogFile"
    Add-LogEntry "ERREUR FATALE: $($_.Exception.Message)"
    $script:ErrorCount++
} finally {
    Write-ColorOutput "White" "`nAppuyez sur Entree pour fermer..."
    Read-Host
}
