text
# Déploiement Automatique Rclone WebDAV sur Windows

## 1. Présentation

Ce guide décrit comment déployer et automatiser le montage d’un remote Rclone WebDAV sous Windows en utilisant un script PowerShell interactif basé sur :
- **winget** pour l’installation de rclone et WinFSP
- **NSSM** pour le service Windows
- Génération automatique de la configuration rclone

---

## 2. Prérequis

- Windows 10/11 x64
- Accès administrateur
- Accès internet
- [winget](https://learn.microsoft.com/fr-fr/windows/package-manager/winget/) installé (sinon, le script gère aussi son installation)

---

## 3. Workflow

1. **Lancer le script**
2. **Installation automatique des dépendances** (rclone, WinFSP, NSSM)
3. **Saisie manuelle ou interactive de la conf**
4. **Génération automatique du fichier `rclone.conf`**
5. **Création et configuration automatique du service Windows (`NSSM`)**
6. **Le remote est monté automatiquement au démarrage**

---

## 4. Utilisation du script

### Lancer le script

1. Ouvrir **PowerShell** en tant qu’administrateur
2. Placer le script (`deploy-rclone-final.ps1`) dans le dossier de votre choix
3. Lancer :
.\deploy-rclone-final.ps1

### Paramètres demandés

- **Nom du remote** (ex: monserveur)
- **URL du serveur WebDAV** (ex: https://domaine.exemple.com/webdav)
- **Nom d’utilisateur WebDAV**
- **Mot de passe WebDAV** (saisie en mode caché)
- **Lettre de lecteur pour le montage** (ex: Z)
- **Type du serveur** (Générique, Nextcloud, Owncloud, Autre)

### Actions réalisées automatiquement

- Installation de rclone (via winget)
- Installation de WinFSP (via winget, avec le package correct `WinFsp.WinFsp`)
- Téléchargement et extraction de NSSM dans `C:\nssm`
- Génération de `C:\rclone\rclone.conf` avec chiffrement du mot de passe
- Création du service `rclone_<nom remote>` qui monte automatiquement le remote sur la lettre choisie
- Activation des logs dans `C:\rclone\rclone.log`
- Paramètres utilisés :
--no-console --vfs-cache-mode full --network-mode --dir-cache-time 10s

---

## 5. Gestion du service

- **Démarrer** :
Start-Service rclone_<nomdu-remote>

- **Arrêter** :
Stop-Service rclone_<nomdu-remote>

- **Status** :
Get-Service rclone_<nomdu-remote>

- **Supprimer** :
C:\nssm\nssm.exe remove rclone_<nomdu-remote>


---

## 6. Personnalisation

- Le script peut être adapté pour déployer plusieurs remotes/services.
- Pour automatisation totale (mode non-interactif), préremplir les variables dans le script ou utiliser un fichier de variables.
- Il est possible de pré-déployer le fichier `rclone.conf` sur plusieurs postes pour réplication de la configuration.

---

## 7. Dépannage

- **Logs** : Analysez `C:\rclone\rclone.log`
- **Montage absent** : Vérifiez WinFSP, lettre de lecteur, droits administrateur
- **Service ne démarre pas** : Vérifiez que les trois briques (nssm, rclone, winfsp) sont bien présentes, le remote accessible, le fichier conf valide

---

## 8. Références

- [Documentation Rclone WebDAV](https://rclone.org/webdav/)
- [WinFSP (Windows FUSE)](https://winfsp.dev/)
- [NSSM, Non Sucking Service Manager](https://nssm.cc/)
- [winget, gestionnaire de paquets Microsoft](https://learn.microsoft.com/fr-fr/windows/package-manager/winget/)

---

> **Astuce** : Vous pouvez organiser chaque section comme page ou sous-page dans un wiki ou outil Markdown structuré (Obsidian, Notion, GitHub Wiki, etc).

Ce script a été créé à l'aide de l'outil Perplexity.
