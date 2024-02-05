<###############################################################################
Script de gestion des utilisateurs AD et Azure
Auteur : Clément Kerviche
Date : 29/01/24
Version : 1.0
###############################################################################>

# -----------------------------------------------------------------------------
# Import des modules
# -----------------------------------------------------------------------------

# Import des modules pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing  

# Import du module AD 
Import-Module ActiveDirectory

# Import du module Exchange Online
Import-Module ExchangeOnlineManagement

# Import du module MGGraph
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Authentication

# -----------------------------------------------------------------------------
# Connexion aux services
# -----------------------------------------------------------------------------

#Connexion à Exchange Online avec l'adresse mail de l'administrateur
$acctName = $env:USERNAME + "@administrateur.fr"
Connect-ExchangeOnline -UserPrincipalName $acctName -ShowProgress $true

#Connexion au module MGGraph
Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -NoWelcome

#Connexion au module AzureAD
Connect-MsolService 


# -----------------------------------------------------------------------------
# Import des bibliothèques
# -----------------------------------------------------------------------------

# Import de la bibliothèque pour personnalisez l'attente pour la fonction "patientez" 

. "C:\script\@@biblio.ps1"
<#
function Patientez {
    param ( [int32]$Sec )

    for (($i = 0); $i -lt $Sec; $i++)
    {
        write-host -nonewline "."
        Start-Sleep 1
    }
}
#>

# Import du menu 

. "$PSScriptRoot\assets\scripts\menu.ps1"

# -----------------------------------------------------------------------------
# Fonctions 
# -----------------------------------------------------------------------------

#fonction qui force une réplication entre tous les contrôleurs de domaine
function Sync-AllDomainController {
    TRY {
        (Get-ADDomainController -Filter *).Name | Foreach-Object { repadmin /syncall $_ (Get-ADDomain).DistinguishedName /e /A | Out-Null };
        patientez 30;
    }
    CATCH {
        Write-Warning "`n`tLa replication automatique n'a pas aboutie, merci de lancer une replication manuelle."
        Pause
        Exit
    }
}

#fonction pour forcer une synchronisation ADconnect 
function ADsynchro {
    TRY {
        $session = New-PSSession -ComputerName BCOADC3.BCO.LOCAL
        Invoke-Command -Session $session -ScriptBlock { Import-Module -Name 'ADSync' }
        Invoke-Command -Session $session -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
        Remove-PSSession $session 
        patientez 40
    }
    CATCH {
        Write-Warning "`n`tLa replication automatique n'a pas aboutie, merci de lancer une replication manuelle."
        While ($sync -ne "o") {
            Write-Host "`nMerci de lancer une replication entre les deux controleurs de domaine et une synchronisation AD Connect" -ForegroundColor Red
            $sync = Read-Host "`nSaisissez 'o' pour continuer"
        }
        Patientez 10
    }
}

#fonction pour gérer les caractères spéciaux
function Remove-StringLatinCharacters {
    PARAM (
        [parameter(ValueFromPipeline = $true)]
        [string]$emailuser
    )
    PROCESS {
        [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($emailuser))
    }
}

#fonction pour gérer les caractères spéciaux dans l'email
function TransformUserEmail {
    PARAM (
        [parameter(ValueFromPipeline = $true)]
        [string]$emailuser
    )
    PROCESS {
        # Mise en place de l'email
        $baseEmail = $firstName.Text.ToLower() + "." + $lastName.Text.ToLower()

        #suppression des tirets dans le champs utilisateur de l'adresse email (partie avant le @)
        $baseEmail = $baseEmail.replace("-", "")
        
        #suppression des caractères spéciaux
        $baseEmail = $baseEmail | Remove-StringLatinCharacters

        #conversion des caractères en minuscule
        $baseEmail = $baseEmail.ToLower()

        #suppression des espaces
        $baseEmail = $baseEmail.replace(" ", "")

        #Préparation du changement de la propriété mail
        $emailuser = $baseEmail + "@laboulangere-co.fr"

        # Vérification qu'il n'y ai pas d'autre utilisateur avec la meme adresse email
        try {
            $existUserAD = Get-ADUser -Filter "mail -eq '$emailuser'" -Properties mail
        }
        catch {
            Write-Warning "`n`tIl n'existe pas d'homonyme dans l'AD"
        }
        
        try {
            $allUsers = Get-MgUser -All:$true -Select UserPrincipalName, Mail
            $existUserAAD = $allUsers | Where-Object { $_.UserPrincipalName -eq $emailuser }
        }
        catch {
            Write-Warning "`n`tIl n'existe pas d'homonyme dans Azure AD"
        }

        $counter = 1
        while ($existUserAD -or $existUserAAD) {
            #Préparation du changement de la propriété mail
            $emailuser = $baseEmail + "$($counter)@laboulangere-co.fr"

            # Vérification qu'il n'y ai pas d'autre utilisateur avec la meme adresse email
            try {
                $existUserAD = Get-ADUser -Filter "mail -eq '$emailuser'" -Properties mail
            }
            catch {
                Write-Warning "`n`tIl n'existe pas d'homonyme dans l'AD"
            }
            
            try {
                $allUsers = Get-MgUser -All:$true -Select UserPrincipalName, Mail
                $existUserAAD = $allUsers | Where-Object { $_.UserPrincipalName -eq $emailuser }
            }
            catch {
                Write-Warning "`n`tIl n'existe pas d'homonyme dans Azure AD"
            }

            $counter++
        }

        return $emailuser
    }
}

#fonction de réinitialisation des champs du formulaire AD
function Reset-Form {
    $firstName.Text = ""
    $lastName.Text = ""
    $entity.SelectedIndex = -1
    $employeID.Text = ""
    $department.Text = ""
    $job.Text = ""
    $manager.Text = ""
    $displayname_label.Text = ""
    $checkbox_cdd.Checked = $false
    $datePicker.Visible = $false
    $datePicker_label.Visible = $false
    $azureAD.Checked = $true
    $E3Label.Text = "E3 disponibles : $NbE3"
    $P1Label.Text = "P1 disponibles : $NbP1"
}

#fonction de réinitialisation des champs du formulaire Azure AD
function Reset-AZForm{
    $firstName.Text = ""
    $lastName.Text = ""
    $entity.SelectedIndex = -1
    $employeID.Text = ""
    $department.Text = ""
    $job.Text = ""
    $manager.Text = ""
}

#--------------------------------------------------------------------------------
# Variables script
#--------------------------------------------------------------------------------

$Script:LogonDC = "BCODC1"

#--------------------------------------------------------------------------------
# Début du script
#--------------------------------------------------------------------------------

Show-Menu