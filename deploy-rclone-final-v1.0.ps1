#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Script de deploiement automatique de rclone avec montage WebDAV
.DESCRIPTION
    Ce script installe automatiquement rclone, WinFSP, NSSM et configure un service de montage WebDAV
.NOTES
    Auteur: Script de deploiement rclone
    Version: 1.4
    Necessite les droits administrateur et winget
#>

# Configuration des couleurs pour l'affichage
$Host.UI.RawUI.WindowTitle = "Deploiement automatique rclone"

function Write-ColorOutput($ForegroundColor, $Message) {
    Write-Host $Message -ForegroundColor $ForegroundColor
}

function Write-Header($Message) {
    Write-Host "`n" + "="*60 -ForegroundColor Yellow
    Write-Host $Message -ForegroundColor Yellow
    Write-Host "="*60 -ForegroundColor Yellow
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

function Install-Winget {
    if (!(Test-Winget)) {
        Write-ColorOutput "Yellow" "Winget n'est pas disponible. Installation en cours..."
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
    
    Write-ColorOutput "Yellow" "Installation de rclone..."
    try {
        $result = winget install Rclone.Rclone --accept-package-agreements --accept-source-agreements --silent 2>&1
        Write-ColorOutput "Green" "rclone installe avec succes."
    } catch {
        Write-ColorOutput "Red" "Erreur lors de l'installation de rclone: $($_.Exception.Message)"
    }
    
    Write-ColorOutput "Yellow" "Installation de WinFSP..."
    try {
        # Le nom correct pour WinFSP est WinFsp.WinFsp
        $result = winget install WinFsp.WinFsp --accept-package-agreements --accept-source-agreements --silent 2>&1
        Write-ColorOutput "Green" "WinFSP installe avec succes."
    } catch {
        Write-ColorOutput "Red" "Erreur lors de l'installation de WinFSP: $($_.Exception.Message)"
    }
    
    Write-ColorOutput "Yellow" "Installation de NSSM..."
    $nssmPath = "C:\nssm"
    if (!(Test-Path $nssmPath)) {
        New-Item -Path $nssmPath -ItemType Directory -Force | Out-Null
        $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
        $nssmZip = "$env:TEMP\nssm.zip"
        
        try {
            Write-ColorOutput "Yellow" "Telechargement de NSSM..."
            $progressPreference = 'silentlyContinue'
            Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip
            
            Write-ColorOutput "Yellow" "Extraction de NSSM..."
            Expand-Archive -Path $nssmZip -DestinationPath $env:TEMP -Force
            
            $nssmSource = "$env:TEMP\nssm-2.24\win64"
            Copy-Item -Path "$nssmSource\*" -Destination $nssmPath -Force
            
            Remove-Item $nssmZip -Force
            Remove-Item "$env:TEMP\nssm-2.24" -Recurse -Force
            
            Write-ColorOutput "Green" "NSSM installe avec succes."
        } catch {
            Write-ColorOutput "Red" "Erreur lors de l'installation de NSSM: $($_.Exception.Message)"
        }
    } else {
        Write-ColorOutput "Green" "NSSM deja installe."
    }
    
    $currentPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
    if ($currentPath -notlike "*$nssmPath*") {
        [Environment]::SetEnvironmentVariable("Path", "$currentPath;$nssmPath", [EnvironmentVariableTarget]::Machine)
        $env:Path += ";$nssmPath"
    }
    
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

function Get-UserConfiguration {
    Write-Header "CONFIGURATION DU REMOTE WEBDAV"
    
    $config = @{}
    
    $config.RemoteName = Read-Host "Nom du remote rclone (ex: monserveur)"
    $config.WebDAVUrl = Read-Host "URL du serveur WebDAV (ex: https://monserveur.com/webdav)"
    $config.Username = Read-Host "Nom d'utilisateur WebDAV"
    
    $securePassword = Read-Host "Mot de passe WebDAV" -AsSecureString
    $config.Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
    
    do {
        $config.DriveLetter = Read-Host "Lettre de lecteur pour le montage (ex: Z)"
        $config.DriveLetter = $config.DriveLetter.ToUpper()
        if ($config.DriveLetter -notmatch "^[D-Z]$") {
            Write-ColorOutput "Red" "Veuillez choisir une lettre entre D et Z."
        }
    } while ($config.DriveLetter -notmatch "^[D-Z]$")
    
    Write-ColorOutput "Cyan" "`nChoisissez le type de serveur WebDAV:"
    Write-ColorOutput "White" "1. Generique (par defaut)"
    Write-ColorOutput "White" "2. Nextcloud"
    Write-ColorOutput "White" "3. Owncloud"
    Write-ColorOutput "White" "4. Autre"
    
    $vendorChoice = Read-Host "Votre choix (1-4)"
    switch ($vendorChoice) {
        "2" { $config.Vendor = "nextcloud" }
        "3" { $config.Vendor = "owncloud" }
        "4" { $config.Vendor = Read-Host "Specifiez le vendor" }
        default { $config.Vendor = "other" }
    }
    
    return $config
}

function Create-RcloneConfig($config) {
    Write-Header "CREATION DE LA CONFIGURATION RCLONE"
    
    $rcloneDir = "C:\rclone"
    $configPath = "$rcloneDir\rclone.conf"
    
    if (!(Test-Path $rcloneDir)) {
        New-Item -Path $rcloneDir -ItemType Directory -Force | Out-Null
        Write-ColorOutput "Green" "Dossier $rcloneDir cree."
    }
    
    Write-ColorOutput "Yellow" "Chiffrement du mot de passe..."
    try {
        $rcloneExe = Get-RcloneExecutable
        $encryptedPassword = & $rcloneExe obscure $config.Password
        
        $configContent = @"
[$($config.RemoteName)]
type = webdav
url = $($config.WebDAVUrl)
vendor = $($config.Vendor)
user = $($config.Username)
pass = $encryptedPassword
"@
    } catch {
        Write-ColorOutput "Yellow" "Impossible de chiffrer le mot de passe, utilisation en clair"
        $configContent = @"
[$($config.RemoteName)]
type = webdav
url = $($config.WebDAVUrl)
vendor = $($config.Vendor)
user = $($config.Username)
pass = $($config.Password)
"@
    }
    
    $configContent | Out-File -FilePath $configPath -Encoding UTF8
    Write-ColorOutput "Green" "Configuration rclone creee: $configPath"
    
    return $configPath
}

function Get-RcloneExecutable {
    $possiblePaths = @(
        "rclone.exe",
        "C:\Program Files\Rclone\rclone.exe",
        "C:\Program Files (x86)\Rclone\rclone.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Rclone.Rclone_*\rclone.exe",
        "C:\Users\$env:USERNAME\AppData\Local\Microsoft\WinGet\Packages\Rclone.Rclone_*\rclone.exe"
    )
    
    foreach ($path in $possiblePaths) {
        if ($path -like "*\*") {
            $expandedPaths = Get-ChildItem $path -ErrorAction SilentlyContinue
            if ($expandedPaths) {
                return $expandedPaths[0].FullName
            }
        } else {
            try {
                $null = Get-Command $path -ErrorAction Stop
                return $path
            } catch { 
                continue
            }
        }
    }
    
    throw "Impossible de trouver rclone.exe"
}

function Create-RcloneService($config, $configPath) {
    Write-Header "CREATION DU SERVICE WINDOWS"
    
    $serviceName = "rclone_$($config.RemoteName)"
    $rcloneExe = Get-RcloneExecutable
    $logPath = "C:\rclone\rclone.log"
    
    $arguments = "mount $($config.RemoteName): $($config.DriveLetter): --no-console --vfs-cache-mode full --network-mode --dir-cache-time 10s --config `"$configPath`" --log-file `"$logPath`" --log-level INFO"
    
    Write-ColorOutput "Yellow" "Suppression du service existant (si present)..."
    & C:\nssm\nssm.exe remove $serviceName confirm 2>&1 | Out-Null
    
    Write-ColorOutput "Yellow" "Creation du nouveau service: $serviceName"
    $installResult = & C:\nssm\nssm.exe install $serviceName "`"$rcloneExe`"" $arguments 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        throw "Erreur lors de la creation du service NSSM: $installResult"
    }
    
    # Configuration du service (rediriger la sortie)
    & C:\nssm\nssm.exe set $serviceName AppDirectory "C:\rclone" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName DisplayName "Rclone Mount - $($config.RemoteName)" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName Description "Service de montage rclone pour $($config.RemoteName)" 2>&1 | Out-Null
    & C:\nssm\nssm.exe set $serviceName Start SERVICE_AUTO_START 2>&1 | Out-Null
    
    Write-ColorOutput "Green" "Service $serviceName cree avec succes."
    
    # Retourner seulement le nom du service
    return $serviceName
}

function Start-RcloneService($serviceName) {
    Write-Header "DEMARRAGE DU SERVICE"
    
    Write-ColorOutput "Yellow" "Demarrage du service $serviceName..."
    try {
        Start-Service $serviceName
        Start-Sleep 5
        
        $service = Get-Service $serviceName
        if ($service.Status -eq "Running") {
            Write-ColorOutput "Green" "Service demarre avec succes!"
        } else {
            Write-ColorOutput "Red" "Erreur lors du demarrage du service."
            Write-ColorOutput "Yellow" "Verifiez le fichier log: C:\rclone\rclone.log"
        }
    } catch {
        Write-ColorOutput "Red" "Erreur lors du demarrage: $($_.Exception.Message)"
        Write-ColorOutput "Yellow" "Verifiez le fichier log: C:\rclone\rclone.log"
    }
}

function Show-Summary($config, $serviceName) {
    Write-Header "RESUME DE L'INSTALLATION"
    
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
    Write-ColorOutput "White" "Log: C:\rclone\rclone.log"
    Write-ColorOutput "White" "Config: C:\rclone\rclone.conf"
    
    Write-ColorOutput "Yellow" "`nPour gerer le service:"
    Write-ColorOutput "White" "Demarrer: Start-Service $serviceName"
    Write-ColorOutput "White" "Arreter: Stop-Service $serviceName"
    Write-ColorOutput "White" "Status: Get-Service $serviceName"
    Write-ColorOutput "White" "Supprimer: C:\nssm\nssm.exe remove $serviceName"
    
    Write-ColorOutput "Green" "`nInstallation terminee avec succes!"
}

# ===== SCRIPT PRINCIPAL =====

try {
    Clear-Host
    Write-ColorOutput "Green" "Script de deploiement automatique rclone WebDAV"
    Write-ColorOutput "Green" "Version 1.4 - Utilise winget"
    
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
    
} catch {
    Write-ColorOutput "Red" "Erreur durant l'installation: $($_.Exception.Message)"
    Write-ColorOutput "Yellow" "Consultez les logs pour plus d'informations."
} finally {
    Write-ColorOutput "White" "`nAppuyez sur Entree pour fermer..."
    Read-Host
}