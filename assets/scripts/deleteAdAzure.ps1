<###############################################################################
Script de suppression d'un utilisateur AD et Azure 
Ce script est appelé par le menu principal du script ManageUser.ps1
Il permet de supprimer un utilisateur AD et Azure avec les possibilités suivantes :
- Archivage de la BAL
- Suppression des licences O365
- Suppression du groupe VPN MFA
###############################################################################>

# --------------------------------------------------------------------------------
# Fenetre du formulaire
# --------------------------------------------------------------------------------

# Création de la fenêtre
$form = New-Object System.Windows.Forms.Form
$form.Text = "La Boulangère - Suppression d'un utilisateur AD et Azure"
$form.Size = New-Object System.Drawing.Size(400,280)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedSingle'

# Mise en place du logo
$icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
$form.Icon = $icon

# Création du label de titre
$title = New-Object System.Windows.Forms.Label
$title.Location = New-Object System.Drawing.Point(20,20)
$title.Size = New-Object System.Drawing.Size(500,20)
$title.Text = "Suppression d'un utilisateur AD et Azure"
$title.Font = New-Object System.Drawing.Font("Tahoma", 12, [System.Drawing.FontStyle]::Bold)

# Création d'un texte explicatif
$text = New-Object System.Windows.Forms.Label
$text.Location = New-Object System.Drawing.Point(20,60)
$text.Size = New-Object System.Drawing.Size(500,20)
$text.Text = "Veuillez saisir le pnom de l'utilisateur à supprimer :"

# Création du champ de saisie du pnom
$pnomDel = New-Object System.Windows.Forms.TextBox
$pnomDel.Location = New-Object System.Drawing.Point(20,100)
$pnomDel.Size = New-Object System.Drawing.Size(200,20)

# Création du bouton de recherche
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Location = New-Object System.Drawing.Point(240,100)
$buttonSearch.Size = New-Object System.Drawing.Size(100,20)
$buttonSearch.BackColor = "#3b82f6"
$buttonSearch.Text = "Rechercher"

# Ajoutez une action sur le clic du bouton de recherche
$buttonSearch.Add_Click({

    # Recherche de l'utilisateur dans AD
    $user = Get-ADUser -Filter {SamAccountName -eq $pnomDel.Text} -Properties * -Server $script:LogonDC

    # Si l'utilisateur n'existe pas
    if ($user -eq $null) {
        # Recherche 2 : on enlève la première lettre et on recherche avec le reste
        $nom = $pnomDel.Text.Substring(1)
        $users = Get-ADUser -Filter {Surname -eq $nom} -Properties * -Server $script:LogonDC

        $usersCount = ($users | Measure-Object).Count
        # Si on trouve plusieurs personnes
        if ($usersCount -gt 1) {
            # On regarde avec la lettre qu'on a enlevé si elle est égale à la première lettre du prénom
            $firstLetter = $pnomDel.Text.Substring(0, 1)
            $users = $users | Where-Object { $_.GivenName.StartsWith($firstLetter) }

            $usersCount = ($users | Measure-Object).Count
            # Si plusieurs utilisateurs sont trouvés
            if ($usersCount -gt 1) {
                $displayname_label.Text = "Plusieurs utilisateurs trouvés"
            }
            # Si un seul utilisateur est trouvé
            elseif ($usersCount -eq 1) {
                $displayname_label.Text = $users[0].DisplayName
                $pnomDel.Text = $users[0].SamAccountName
            }
            # Si aucun utilisateur n'est trouvé
            else {
                $displayname_label.Text = "Utilisateur non trouvé"
            }
        } elseif ($usersCount -eq 1) {
            $displayname_label.Text = $users[0].DisplayName
            $pnomDel.Text = $users[0].SamAccountName
        } else {
            $displayname_label.Text = "Utilisateur non trouvé"
        }
    }
    # Sinon on affiche le DisplayName de l'utilisateur
    else {
        $displayname_label.Text = $user.DisplayName
    }
})

# Label pour afficher le DisplayName du responsable
$displayname_label = New-Object System.Windows.Forms.Label
$displayname_label.Location = New-Object System.Drawing.Size(20, 120)
$displayname_label.Size = New-Object System.Drawing.Size(100, 40)

# Checkbox d'archibage de la BAL
$checkboxArchive = New-Object System.Windows.Forms.CheckBox
$checkboxArchive.Location = New-Object System.Drawing.Point(240,140)
$checkboxArchive.Size = New-Object System.Drawing.Size(200,20)
$checkboxArchive.Checked = $true
$checkboxArchive.Text = "Archiver la BAL"

# Bouton d'anulation
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(20,180)
$buttonCancel.Size = New-Object System.Drawing.Size(100,40)
$buttonCancel.BackColor = "#ef4444"
$buttonCancel.Text = "Annuler"

# Action du bouton d'annulation retour au menu
$buttonCancel.Add_Click({
    $form.Visible = $false
    Show-Menu
})

# Bouton de validation
$buttonDel = New-Object System.Windows.Forms.Button
$buttonDel.Location = New-Object System.Drawing.Point(240,180)
$buttonDel.Size = New-Object System.Drawing.Size(100,40)
$buttonDel.BackColor = "#22c55e"
$buttonDel.Text = "Supprimer"

# --------------------------------------------------------------------------------
# Action du bouton de validation
# --------------------------------------------------------------------------------
$buttonDel.Add_Click({ 

    # Recherche de l'utilisateur dans AD et récuperation de ses propriétés
    $user = Get-ADUser -Filter {SamAccountName -eq $pnomDel.Text} -Properties * -Server $script:LogonDC

    # Récupération de la date du jour
    $today= Get-Date -Format "dd/MM/yyyy"

    # Si l'utilisateur existe
    if($user){
        # Désactivation du compte
        Write-Host "`nDésactivation du compte..." -ForegroundColor Yellow
        TRY {
            Disable-ADAccount -Server $Script:LogonDC -Identity $pnomDel.Text
            
            # Préparation pour le renseignement de la propriété description
            $newdesc = $user.Description + " - Désactivé le $today ($env:username)"           
            Set-aduser -Identity $pnomDel.Text -Description $newdesc -Server $Script:LogonDC

            # Préparation pour le renseignement de la propriété ExtensionAttribute10
            $extensionAttribute10 = $user
            $extensionAttribute10."extensionAttribute10" = $today
            Set-ADUser -Instance $extensionAttribute10 -Server $Script:LogonDC

            Write-Host "`tLe compte a été désactivé" -ForegroundColor Green
        }
        CATCH {
            Write-Warning "`tLe compte n'a pas pu être désactivé" -ForegroundColor Red
        }

        # Purge du champ responsable 
        TRY {
            Set-ADUser -identity $pnomDel.Text -manager $null -Server $Script:LogonDC
            Write-Host "`tLe champ 'responsable' a été purgé." -ForegroundColor Green
        }
        CATCH {
            Write-Warning "`tImpossible de purger le champ 'Responsable' " -ForegroundColor Red
        }

        # Suppression du groupe VPN MFA (libération license INWEBO)
        if ($user.MemberOf -match "CN=DL-inwebo-Acces_VPN_MFA_DSI") {
            TRY {
                Remove-AdGroupMember -Identity DL-inwebo-Acces_VPN_MFA_DSI -Members $pnomDel.Text -Confirm:$false -Server $Script:LogonDC
                Write-Host "`tL'utilisateur a été retiré du groupe DL-inwebo-Acces_VPN_MFA_DSI" -ForegroundColor Green
            }
            CATCH {
                Write-Warning "`tL'utilisateur n'a pu être retiré du groupe DL-inwebo-Acces_VPN_MFA_DSI" -ForegroundColor Red
            }
        }
        elseif ($user.MemberOf -match "CN=DL-inwebo-Acces_VPN_MFA") {
            TRY {
                Remove-AdGroupMember -Identity DL-inwebo-Acces_VPN_MFA -Members $pnomDel.Text -Confirm:$false -Server $Script:LogonDC
                Write-Host "`tL'utilisateur a été retiré du groupe DL-inwebo-Acces_VPN_MFA" -ForegroundColor Green
            }
            CATCH {
                Write-Warning "`tL'utilisateur n'a pu être retiré du groupe DL-inwebo-Acces_VPN_MFA" -ForegroundColor Red
            }
        }

        # Déplacement du compte désactivé dans l'OU "OU=COMPTES DESACTIVES,OU=BCO,DC=BCO,DC=local"
        TRY {
            Move-ADObject -Identity  $user.DistinguishedName -TargetPath "OU=COMPTES DESACTIVES,OU=BCO,DC=BCO,DC=local" -Server $Script:LogonDC
            Write-Host "`tLe compte a été déplacé dans l'OU 'OU=COMPTES DESACTIVES,OU=BCO,DC=BCO,DC=local'" -ForegroundColor Green
        }
        CATCH {
            Write-Warning "`tLe compte n'a pas pu être déplacé dans l'OU 'OU=COMPTES DESACTIVES,OU=BCO,DC=BCO,DC=local'" -ForegroundColor Red
        }

        # Suppression de la propriété msExchExtensionCustomAttribute1: (filtrage AD Connect)
        if ($user.msExchExtensionCustomAttribute1 -eq "O365") {
            TRY {
                Set-ADUser -Identity $pnomDel.Text -clear msExchExtensionCustomAttribute1 -Server $Script:LogonDC
                Write-Host "`tL'attribut msExchExtensionCustomAttribute1 a bien été supprimé" -ForegroundColor Green
            }
            CATCH {
                Write-Warning "`tL'attribut msExchExtensionCustomAttribute1 n'a pas pu être supprimé" -ForegroundColor Red
            } 
        }
        else {
            Write-Host "`tL'attribut msExchExtensionCustomAttribute1 n'est pas renseigné" -ForegroundColor Red
        }

        # Suppression des licences O365 si l'utilisateur n'est pas archivé
        if (-not $checkboxArchive.Checked){
            # Verification sur l'utilisateur a des licences O365
            $userLicenses = Get-MgUserLicenseDetail -UserId $user.UserPrincipalName

            #Si il a des licences alors on les supprimes
            if ($userLicenses) {
                TRY {
                    # Réccupération des licences de l'utilisateur
                    $userLicenses = (Get-MgUserLicenseDetail -UserId $user.UserPrincipalName).SkuId

                    # Suppression des licences de l'utilisateur
                    foreach ($license in $userLicenses) {
                        Set-MgUserLicense -UserId $user.UserPrincipalName -RemoveLicenses @($license) -AddLicenses @()
                    }
                    Write-Host "`tLes licences O365 ont été supprimées" -ForegroundColor Green
                }
                CATCH {
                    Write-Warning "`tLes licences O365 n'ont pas pu être supprimées" -ForegroundColor Red
                } 
            }
            else {
                Write-Host "`tL'utilisateur n'a pas de licence O365" -ForegroundColor Red
            }
            [System.Windows.Forms.MessageBox]::Show("L'utilisateur $($user.DisplayName) à été :`n`t-Désactivé`n`t-Supprimé des licences O365", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        } 
    else {
        Write-Warning "`tL'utilisateur n'a pas pu être trouvé" -ForegroundColor Red
    }

    # Lancement d'une replication entre les DC BCO
    Write-Host "`nLancement d'une replication Active Directory" -ForegroundColor Yellow
    Sync-AllDomainController

    # Lancement d'une synchronisation AD Connect
    Write-Host "`nLancement d'une synchronisation AD Connect..." -ForegroundColor Yellow
    ADsynchro

    # Si archivage de la bal on restaure l'utilisateur coté azure
    if ($checkboxArchive.Checked){
        try {
            # Recherche du manager de l'utilisateur
            $manager = Get-ADUser -Identity $user.Manager -Properties * -Server $script:LogonDC

            # Initialisez deletedUser à null
            $deletedUser = $null

            $seconds = 0
            # Tant que l'utilisateur n'est pas trouvé dans la corbeille de recyclage
            while ($null -eq $deletedUser) {
                # Obtenez tous les utilisateurs supprimés
                $deletedUsers = Get-MsolUser -ReturnDeletedUsers

                # Trouvez l'utilisateur que vous voulez restaurer
                $deletedUser = $deletedUsers | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }

                if ($null -eq $deletedUser) {
                    Write-Host "`n`tL'utilisateur n'existe pas dans la corbeille de recyclage. Nouvelle vérification dans 15 secondes..."
                    patientez 15
                    $seconds += 15
                }
            }

            Write-Host "`n`tL'utilisateur a été trouvé dans la corbeille de recyclage après $seconds secondes." -ForegroundColor Green

            # Restaurez l'utilisateur
            Write-Host "`nAttente de la restauration de l'utilisateur..." -ForegroundColor Yellow
            ADsynchro
            Write-Host "`nRestauration en cours..." -ForegroundColor Yellow
            ADsynchro

            try {
                Restore-MsolUser -UserPrincipalName $deletedUser.UserPrincipalName
                Write-Host "`n`tL'utilisateur a été restauré" -ForegroundColor Green
            }
            catch {
                Write-Host "`n`tL'utilisateur n'a pas pu être restauré" -ForegroundColor Red
                Write-Host "`n`tMerci de le restaurer manuellement dans la corbeille de recyclage Azure AD" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "`n`tUne erreur s'est produite lors de la recherche de l'utilisateur" -ForegroundColor Red
        }

            # Attendez que l'utilisateur soit restauré
            Write-Host "`nAttente de la restauration totale de l'utilisateur..." -ForegroundColor Yellow
            patientez 30

            # Initialisation de $userRestore à $null
            $userRestore = $null
            # Initialisation du compteur de temps d'attente
            $seconds = 0
            # Boucle tant que $userRestore est $null (c'est-à-dire, l'utilisateur n'existe pas)
            while ($null -eq $userRestore) {
                $userRestore = Get-MgUser -All | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
                if ($null -ne $userRestore) {
                    # L'utilisateur a été trouvé
                    Write-Host "`nL'utilisateur $($userRestore.DisplayName) a été trouvé en $($seconds) secondes" -ForegroundColor Green
                    # On renomme avec [Archive] devant le nom de l'utilisateur
                    $newDisplayName = "[Archive] $($userRestore.DisplayName)"
                    Update-MgUser -UserId $userRestore.Id -DisplayName $newDisplayName
                    Write-Host "`n`tLe nom de l'utilisateur a été modifié" -ForegroundColor Green

                    # On modifie l'état de connexion de l'utilisateur en bloqué
                    try {
                        Update-MgUser -UserId $userRestore.Id -AccountEnabled:$false
                        Write-Host "`n`tL'état de connexion de l'utilisateur a été modifié" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "`n`tL'état de connexion de l'utilisateur n'a pas pu être modifié" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }
                }
                else {
                    # L'utilisateur n'existe pas, attendre 15 secondes avant de réessayer
                    Write-Host "`n`tL'utilisateur n'existe pas. Nouvelle vérification dans 15 secondes..."
                    patientez 15
                    # Augmenter le compteur de temps d'attente
                    $seconds += 15
                }
            }

        # On masque l'adresse mail du carnet d'adresse 
        $mailbox = $null
        $seconds = 0
        $attempts = 0
        while ($null -eq $mailbox) {
            $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($null -ne $mailbox) {
                do {
                    try {
                        Set-Mailbox -Identity $user.UserPrincipalName -HiddenFromAddressListsEnabled:$true
                        Write-Host "`n`tL'adresse mail de l'utilisateur a été masquée après $seconds secondes." -ForegroundColor Green
                        $success = $true
                    }
                    catch {
                        $attempts++
                        if ($attempts -ge 4) {
                            Write-Host "Impossible de masquer l'adresse mail de l'utilisateur, merci de le faire via le portail O365." -ForegroundColor Red
                            break
                        }
                        Write-Host "Une erreur s'est produite. Nouvelle tentative dans 15 secondes..."
                        patientez 15
                        $success = $false
                    }
                } while (!$success)
            }
            else {
                Write-Host "`nLa boîte aux lettres de l'utilisateur n'est pas encore remontée. Nouvelle vérification dans 15 secondes..."
                patientez 15
                $seconds += 15
            }
        }

        # On passe le type de boite aux lettres en boite aux lettres partagée

        try {
            Set-Mailbox -Identity $user.UserPrincipalName -Type Shared
            Write-Host "`n`tLa boîte aux lettres de l'utilisateur a été convertie en boîte aux lettres partagée" -ForegroundColor Green
        }
        catch {
            Write-Host "`n`tLa boîte aux lettres de l'utilisateur n'a pas pu être convertie en boîte aux lettres partagée" -ForegroundColor Red
        }

        # On met en délégation son manager en accès total à la boite mail
        try {
            Add-MailboxPermission -Identity $user.UserPrincipalName -User $manager.UserPrincipalName -AccessRights FullAccess -InheritanceType All
            Write-Host "`n`tL'utilisateur $($user.DisplayName) a été ajouté en délégation à son manager $($manager.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "`n`tL'utilisateur n'a pas pu être ajouté en délégation à son manager" -ForegroundColor Red
        }

        # On informe l'utilisateur du reste de la procédure a effectué
        [System.Windows.Forms.MessageBox]::Show("L'utilisateur $($user.DisplayName) à été :`n`t-Désactivé`n`t-Archivée`n`t-Convertit en BAL partagée`n`t-Mis en délégation totale : $($manager.DisplayName)`nMerci d'attendre 24h avant d'enlever les licences, et vérifié les délégations si besoins.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }

    # Fin du traitement
    $result = [System.Windows.Forms.MessageBox]::Show("Le traitement est terminé.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    if ($result -eq "OK") {
        $form.Visible = $false
        Show-Menu
    }
})


# --------------------------------------------------------------------------------
# Affichage du formulaire
# --------------------------------------------------------------------------------

$form.Controls.Add($title)
$form.Controls.Add($text)
$form.Controls.Add($pnomDel)
$form.Controls.Add($buttonSearch)
$form.Controls.Add($displayname_label)
$form.Controls.Add($checkboxArchive)
$form.Controls.Add($buttonCancel)
$form.Controls.Add($buttonDel)

$form.ShowDialog() > $null
