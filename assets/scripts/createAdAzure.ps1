<###############################################################################
Script de création de compte AD et boîte mail
Ce script est appelé par le menu principal du script ManageUser.ps1
Il permet de Crée un utilisateur AD et Azure ou uniquement AD 
###############################################################################>

#--------------------------------------------------------------------------------
# Fonctions
#--------------------------------------------------------------------------------

# Fonction qui permet de mettre en place le mot de passe personnalisé pour les forces de ventes
function FDVSetPassword ($samAccountName) {
    # Si le bouton d'annulation est cliqué, fermez la fenetre
    $FDV_Cancel.Add_Click({
            $FDV.Close()
            return
        })

    # Si le bouton de validation est cliqué, modifiez le mot de passe et fermez la fenetre
    $FDV_Validate.Add_Click({
        try {
            $newUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName} -Server $Script:LogonDC -Properties *
            Set-ADAccountPassword -Identity $newUser.SamAccountName -Reset -NewPassword (ConvertTo-SecureString "password$($FDV_secteur.Text)" -AsPlainText -Force) -Server $Script:LogonDC
            Write-Host "`n`tLe mot de passe a été modifié avec succès" -ForegroundColor Green
        }
        catch {
            Write-Warning $_.Exception.Message
        }
            $FDV.Close()
            return
        })

    # Affichage de la fenetre forces de ventes
    $FDV.Controls.Add($FDV_secteurLabel)
    $FDV.Controls.Add($FDV_secteur)
    $FDV.Controls.Add($FDV_Cancel)
    $FDV.Controls.Add($FDV_Validate)
    $FDV.ShowDialog() > $null
}

#--------------------------------------------------------------------------------
# Définition des variables globales
#--------------------------------------------------------------------------------

# Unité d'organisation par défaut
$ouUser = ",OU=USERS,OU=TEST,DC=EXEMPLE,DC=local"
#$ouUserTest = "OU=UTILISATEURS TEST,OU=TESTS,OU=TEST,DC=EXEMPLE,DC=local"

# Définition du chemin d'accès au fichier jobs.txt
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$jobsFilePath = Join-Path -Path $scriptPath -ChildPath "../jobs.txt"

#Récupération des informations sur les licences
$E3 = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'ENTERPRISEPACK' 
$E3Details = Get-MgSubscribedSKU -All -Property @("SkuId", "SkuPartNumber", "ConsumedUnits", "PrepaidUnits") | 
Where-Object SkuPartNumber -eq 'ENTERPRISEPACK'  | 
Select-Object *, @{Name = "ActiveUnits"; Expression = { ($_ | Select-Object -ExpandProperty PrepaidUnits).Enabled } } | 
Select-Object SkuId, SkuPartNumber, ActiveUnits, ConsumedUnits

$NbActiveE3 = $E3Details.ActiveUnits
$NbConsumeE3 = $E3Details.ConsumedUnits

$NbE3 = $NbActiveE3 - $NbConsumeE3

$P1 = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'AAD_PREMIUM' 
$P1Details = Get-MgSubscribedSKU -All -Property @("SkuId", "SkuPartNumber", "ConsumedUnits", "PrepaidUnits") | 
Where-Object SkuPartNumber -eq 'AAD_PREMIUM'  | 
Select-Object *, @{Name = "ActiveUnits"; Expression = { ($_ | Select-Object -ExpandProperty PrepaidUnits).Enabled } } | 
Select-Object SkuId, SkuPartNumber, ActiveUnits, ConsumedUnits

$NbActiveP1 = $P1Details.ActiveUnits
$NbConsumeP1 = $P1Details.ConsumedUnits

$NbP1 = $NbActiveP1 - $NbConsumeP1

#Définition des services à désactiver pour les licences E3 et P1
$addE3Licenses = @(
    @{
        SkuId         = $E3.SkuId
        DisabledPlans = $E3.ServicePlans |
        Where-Object ServicePlanName -in ( "MIP_S_CLP1", "MICROSOFTBOOKINGS", "Deskless", "VIVA_LEARNING_SEEDED", "KAIZALA_O365_P3", "MYANALYTICS_P2", "SWAY", "INTUNE_O365", "YAMMER_ENTERPRISE", "RMS_S_ENTERPRISE", "FLOW_O365_P2", "PROJECTWORKMANAGEMENT") | `
            Select-Object -ExpandProperty ServicePlanId
    }
)

$addP1Licenses = @(
    @{
        SkuId         = $P1.SkuId
        DisabledPlans = $P1.ServicePlans |
        Where-Object ServicePlanName -in ("EXCHANGE_S_FOUNDATION", "ADALLOM_S_DISCOVERY") |
        Select-Object -ExpandProperty ServicePlanId
    }
)

#--------------------------------------------------------------------------------
# Création de la fenêtre
#--------------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form 
$form.Text = "La Boulangère - Création utilisateur"
$form.Size = New-Object System.Drawing.Size(520, 350)
$form.StartPosition = "CenterScreen" 
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

$icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
$form.Icon = $icon

#--------------------------------------------------------------------------------
# Création du visuel du formulaire
#--------------------------------------------------------------------------------

# Label du prénom
$firstName_label = New-Object System.Windows.Forms.Label
$firstName_label.Location = New-Object System.Drawing.Size(20, 10)
$firstName_label.Size = New-Object System.Drawing.Size(100, 20)
$firstName_label.Text = "Prénom :"

# TextBox du prénom
$firstName = New-Object System.Windows.Forms.TextBox
$firstName.Location = New-Object System.Drawing.Size(20, 30)
$firstName.Size = New-Object System.Drawing.Size(100, 20)

# Label du nom
$lastName_label = New-Object System.Windows.Forms.Label
$lastName_label.Location = New-Object System.Drawing.Size(140, 10)
$lastName_label.Size = New-Object System.Drawing.Size(100, 20)
$lastName_label.Text = "Nom :"

# TextBox du nom
$lastName = New-Object System.Windows.Forms.TextBox
$lastName.Location = New-Object System.Drawing.Size(140, 30)
$lastName.Size = New-Object System.Drawing.Size(100, 20)

# Label de l'entité
$entity_label = New-Object System.Windows.Forms.Label
$entity_label.Location = New-Object System.Drawing.Size(260, 10)
$entity_label.Size = New-Object System.Drawing.Size(100, 20)
$entity_label.Text = "Entité :"

# ComboBox (liste déroulante) de l'entité
$entity = New-Object System.Windows.Forms.ComboBox
$entity.Location = New-Object System.Drawing.Size(260, 30)
$entity.Size = New-Object System.Drawing.Size(100, 20)
$entity.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Ajout des éléments dans la liste déroulante des entités
$entity.Items.Add("Beaune Brioche") > $null
$entity.Items.Add("La Boulangere & Co")> $null
$entity.Items.Add("Nor'Pain") > $null
$entity.Items.Add("Ouest Boulangere")> $null
$entity.Items.Add("Pain Concept") > $null
$entity.Items.Add("Painorient")> $null
$entity.Items.Add("U7") > $null
$entity.Items.Add("Viennoiserie Ligerienne") > $null

# Définition des variables globales par rapport à la sélection de l'entité
$entity.Add_SelectedIndexChanged({
    switch ($entity.SelectedItem) {
        "Beaune Brioche" { $global:OU = "BEAUNE" ; $global:societe = "Beaune Brioche"; $global:StreetAddress = "Les Cerisières"; $global:State = "Côte-d'Or"; $global:PostalCode = "21209"; $global:POBox = "BP 357"; $global:City = "Beaune"; Break }
        "La Boulangere & Co" { $global:OU = "LES ESSARTS" ; $global:societe = "La Boulangère & Co"; $global:StreetAddress = "PA La Mongie - 1 rue du Petit Bocage"; $global:State = "Vendée"; $global:PostalCode = "85140"; $global:POBox = "CS 40201"; $global:City = "Essarts-en-Bocage"; Break }
        "Nor'Pain" { $global:OU = "VAL DE SAANE" ; $global:societe = "Nor'Pain"; $global:StreetAddress = "Le Vieux Moulin"; $global:State = "Seine-Maritime"; $global:PostalCode = "76890"; $global:POBox = "BP 70116"; $global:City = "Val de Saâne"; Break }
        "Ouest Boulangere" { $global:OU = "LES HERBIERS" ; $global:societe = "Ouest Boulangère"; $global:StreetAddress = "ZA La Buzenière - 10, rue Olivier de Serres"; $global:State = "Vendée"; $global:PostalCode = "85503"; $global:POBox = "BP 60327"; $global:City = "Les Herbiers"; Break }
        "Pain Concept" { $global:OU = "STE HERMINE" ; $global:societe = "Pain Concept"; $global:StreetAddress = "Parc Atlantique"; $global:State = "Vendée"; $global:PostalCode = "85210"; $global:City = "Sainte Hermine"; Break }
        "Panorient" { $global:OU = "GRETZ" ; $global:societe = "Panorient"; $global:StreetAddress = "ZA Eiffel - 36 rue Eiffel"; $global:State = "Seine-et-Marne"; $global:PostalCode = "77220"; $global:City = "Gretz Armainvilliers"; Break }
        "Viennoiserie Ligerienne" { $global:OU = "MORTAGNE" ; $global:societe = "Viennoiserie Ligérienne"; $global:StreetAddress = "ZI du Gautreau II - 647 rue Antoine Carême"; $global:State = "Vendée"; $global:PostalCode = "85290"; $global:POBox = "BP 60"; $global:City = "Beaune"; Break }
        "U7" { $global:OU = "LA CHAIZE" ; $global:societe = "U7"; $global:StreetAddress = "ZA La Folie - 38 rue Charles Tellier"; $global:State = "Vendée"; $global:PostalCode = "85036"; $global:POBox = "CS 80310"; $global:City = "La Chaize-Le-Vicomte"; Break }
    }
})

# Label de l'ID employé
$employeID_label = New-Object System.Windows.Forms.Label
$employeID_label.Location = New-Object System.Drawing.Size(380, 10)
$employeID_label.Size = New-Object System.Drawing.Size(100, 20)
$employeID_label.Text = "Matricule :"

# TextBox de l'ID employé
$employeID = New-Object System.Windows.Forms.TextBox
$employeID.Location = New-Object System.Drawing.Size(380, 30)
$employeID.Size = New-Object System.Drawing.Size(100, 20)
$employeID.MaxLength = 10

# Label du service
$department_label = New-Object System.Windows.Forms.Label
$department_label.Location = New-Object System.Drawing.Size(20, 60)
$department_label.Size = New-Object System.Drawing.Size(100, 20)
$department_label.Text = "Service :"

# ComboBox (liste déroulante) du service
$department = New-Object System.Windows.Forms.ComboBox
$department.Location = New-Object System.Drawing.Size(20, 80)
$department.Size = New-Object System.Drawing.Size(460, 20)

# Ajout des éléments dans la liste déroulante + autocomplétion
$department.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$department.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$department.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

# Lire le fichier et remplir la liste déroulante des départements
Get-Content -LiteralPath $jobsFilePath | ForEach-Object {
    $departmentName = $_.Split(',')[0]
    if ($department.Items.IndexOf($departmentName) -eq -1) {
        $department.Items.Add($departmentName) > $null
    }
}

# Label de la fonction
$job_label = New-Object System.Windows.Forms.Label
$job_label.Location = New-Object System.Drawing.Size(20, 110)
$job_label.Size = New-Object System.Drawing.Size(100, 20)
$job_label.Text = "Fonction :"

# TextBox de la fonction
$job = New-Object System.Windows.Forms.ComboBox
$job.Location = New-Object System.Drawing.Size(20, 130)
$job.Size = New-Object System.Drawing.Size(460, 20)

# Ajout des éléments dans la liste déroulante + autocomplétion
$job.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$job.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$job.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

# Mettre à jour la liste des postes lorsque le département change
$department.add_SelectedIndexChanged({
    $job.Items.Clear()
    $selectedDepartment = $department.SelectedItem
    Get-Content -LiteralPath $jobsFilePath | Where-Object { $_.StartsWith($selectedDepartment + ',') } | ForEach-Object {
        $jobName = $_.Split(',')[1]
        $job.Items.Add($jobName) > $null
    }
})

# Label du responsable
$manager_label = New-Object System.Windows.Forms.Label
$manager_label.Location = New-Object System.Drawing.Size(20, 160)
$manager_label.Size = New-Object System.Drawing.Size(120, 20)
$manager_label.Text = "Responsable (pnom) :"

# TextBox du responsable
$manager = New-Object System.Windows.Forms.TextBox
$manager.Location = New-Object System.Drawing.Size(20, 180)
$manager.Size = New-Object System.Drawing.Size(100, 20)

# Bouton de recherche du responsable
$button_search = New-Object System.Windows.Forms.Button
$button_search.Location = New-Object System.Drawing.Size(120, 180) 
$button_search.Size = New-Object System.Drawing.Size(20, 20) 
$button_search.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#06b6d4")
$button_search.Text = "R" 

# Ajoutez une action sur le clic du bouton de recherche
$button_search.Add_Click({

    # Obtenez l'utilisateur de l'AD
    $user = Get-ADUser -Filter "SamAccountName -eq '$($manager.Text)'" -Properties * -Server $Script:LogonDC

    # Si l'utilisateur n'existe pas
    if ($user -eq $null) {
        # Recherche 2 : on enlève la première lettre et on recherche avec le reste
        $nom = $manager.Text.Substring(1)
        $users = Get-ADUser -Filter "Surname -eq '$nom'" -Properties * -Server $Script:LogonDC

        $usersCount = ($users | Measure-Object).Count
        # Si on trouve plusieurs personnes
        if ($usersCount -gt 1) {
            # On regarde avec la lettre qu'on a enlevé si elle est égale à la première lettre du prénom
            $firstLetter = $manager.Text.Substring(0, 1)
            $users = $users | Where-Object { $_.GivenName.StartsWith($firstLetter) }

            $usersCount = ($users | Measure-Object).Count
            # Si plusieurs utilisateurs sont trouvés
            if ($usersCount -gt 1) {
                $displayname_label.Text = "Plusieurs utilisateurs trouvés"
            }
            # Si un seul utilisateur est trouvé
            elseif ($usersCount -eq 1) {
                $displayname_label.Text = $users[0].DisplayName
                $manager.Text = $users[0].SamAccountName
            }
            # Si aucun utilisateur n'est trouvé
            else {
                $displayname_label.Text = "Utilisateur non trouvé"
            }
        } elseif ($usersCount -eq 1) {
            $displayname_label.Text = $users.DisplayName
            $manager.Text = $users.SamAccountName
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
$displayname_label.Location = New-Object System.Drawing.Size(20, 200)
$displayname_label.Size = New-Object System.Drawing.Size(100, 40)

# Checkbox pour type de contrat
$checkbox_cdd = New-Object System.Windows.Forms.CheckBox
$checkbox_cdd.Text = "CDD"
$checkbox_cdd.Location = New-Object System.Drawing.Size(20, 260)
$checkbox_cdd.Size = New-Object System.Drawing.Size(40, 20)
$checkbox_cdd.AutoSize = $true

# Création du champ de date de fin de contrat
$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Location = New-Object System.Drawing.Size(80, 260)
$datePicker.Size = New-Object System.Drawing.Size(100, 20)
$datePicker.Visible = $false

# Label du champ de date de fin de contrat
$datePicker_label = New-Object System.Windows.Forms.Label
$datePicker_label.Location = New-Object System.Drawing.Size(80, 240)
$datePicker_label.Size = New-Object System.Drawing.Size(100, 20)
$datePicker_label.Text = "Date de fin :"
$datePicker_label.Visible = $false

# Action lorsque la checkbox est cochée
$checkbox_cdd.Add_Click({
    if ($checkbox_cdd.Checked) {
        $datePicker.Visible = $true
        $datePicker_label.Visible = $true
    }
    else {
        $datePicker.Visible = $false
        $datePicker_label.Visible = $false
    }
})

# Checkbox création azure ad
$azureAD = New-Object System.Windows.Forms.CheckBox
$azureAD.Location = New-Object System.Drawing.Size(180, 180)
$azureAD.Size = New-Object System.Drawing.Size(120, 20)
$azureAD.Checked = $true
$azureAD.Text = "Création boîte mail"

# Label pour afficher le nombre de licences E3 disponibles
$E3Label = New-Object System.Windows.Forms.Label
$E3Label.Location = New-Object System.Drawing.Size(320, 180)
$E3Label.Size = New-Object System.Drawing.Size(200, 20)
$E3Label.Text = "E3 disponibles : $NbE3"

# Label pour afficher le nombre de licences P1 disponibles
$P1Label = New-Object System.Windows.Forms.Label
$P1Label.Location = New-Object System.Drawing.Size(320, 200)
$P1Label.Size = New-Object System.Drawing.Size(200, 20)
$P1Label.Text = "P1 disponibles : $NbP1"

# Fenetre si l'entité est forces de ventes
$FDV = New-Object System.Windows.Forms.Form
$FDV.Text = "La Boulangère - Forces de ventes"
$FDV.Size = New-Object System.Drawing.Size(280, 160)
$FDV.StartPosition = "CenterScreen"
$FDV.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Icone de la fenetre forces de ventes
$icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
$FDV.Icon = $icon

# Label de la fenetre forces de ventes
$FDV_secteurLabel = New-Object System.Windows.Forms.Label
$FDV_secteurLabel.Location = New-Object System.Drawing.Size(20, 10)
$FDV_secteurLabel.Size = New-Object System.Drawing.Size(100, 20)
$FDV_secteurLabel.Text = "Secteur :"

# TextBox de la fenetre forces de ventes pour le secteur
$FDV_secteur = New-Object System.Windows.Forms.TextBox
$FDV_secteur.Location = New-Object System.Drawing.Size(20, 30)
$FDV_secteur.Size = New-Object System.Drawing.Size(100, 20)
$FDV_secteur.MaxLength = 4

# Bouton d'annulation fenetre forces de ventes
$FDV_Cancel = New-Object System.Windows.Forms.Button
$FDV_Cancel.Location = New-Object System.Drawing.Size(20, 60)
$FDV_Cancel.Size = New-Object System.Drawing.Size(100, 40)
$FDV_Cancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ef4444")
$FDV_Cancel.Text = "Annuler"

# Bouton de validation fenetre forces de ventes
$FDV_Validate = New-Object System.Windows.Forms.Button
$FDV_Validate.Location = New-Object System.Drawing.Size(140, 60)
$FDV_Validate.Size = New-Object System.Drawing.Size(100, 40)
$FDV_Validate.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#22c55e")
$FDV_Validate.Text = "Valider"

# Bouton d'annulation
$button_cancel = New-Object System.Windows.Forms.Button
$button_cancel.Location = New-Object System.Drawing.Size(270, 260)
$button_cancel.Size = New-Object System.Drawing.Size(100, 40)
$button_cancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ef4444")
$button_cancel.Text = "Annuler"

# Action du bouton d'annulation
$button_cancel.Add_Click({
    $form.Visible = $false
    Show-Menu
})

# Bouton de validation
$button_validate = New-Object System.Windows.Forms.Button
$button_validate.Location = New-Object System.Drawing.Size(380, 260)
$button_validate.Size = New-Object System.Drawing.Size(100, 40)
$button_validate.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#22c55e")
$button_validate.Text = "Valider"

# Action du bouton de validation 
$button_validate.Add_Click({
# --------------------------------------------------------------------------------
# Début du traitement du formulaire
# --------------------------------------------------------------------------------

        $errorCount = 0

        # Vérification de la saisie du formulaire
        if ($firstName.Text -eq "" -or $lastName.Text -eq "" -or $entity.SelectedItem -eq "" -or $department.Text -eq "" -or $job.Text -eq "" -or $manager.Text -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Veuillez saisir tout les champs", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        #Verification de la saisie du matricule
        if ($employeID.Text -eq ""){
        $result = [System.Windows.Forms.MessageBox]::Show("Le matricule n'est pas renseigné est-ce normal ?`nSi oui alors le matricule prendra comme valeur `"AAAAAAAAAA`".", "Attention", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

            if ($result -eq "No") {
                return
            } else{
                $employeID.Text = "AAAAAAAAAA"
            }
        }

        # Transformation du nom de famille en majuscule
        $lastName.Text = $lastName.Text.ToUpper()

        # Initialisation des variables
        $prenom = $firstName.Text
        $nom = $lastName.Text
        $index = 1

        # Création du sAMAccountName initial
        $samAccountName = $prenom.Substring(0, 1).ToLower() + $nom.ToLower()

        # Vérification de l'unicité du sAMAccountName
        $existingUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -Server $Script:LogonDC -Properties DisplayName
        while ($existingUser -and ![string]::IsNullOrEmpty($existingUser.DisplayName)) {

            # Si le sAMAccountName existe déjà, ajoutez une autre lettre du prénom
            $index++
            $samAccountName = $prenom.Substring(0, $index).ToLower() + $nom.ToLower()

            # Rechercher à nouveau l'utilisateur
            $existingUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -Server $Script:LogonDC -Properties DisplayName
        }

        #Informer que le samaccountname est différent de la normalisation
        if ($samAccountName -ne $prenom.Substring(0, 1).ToLower() + $nom.ToLower()) {
            [System.Windows.Forms.MessageBox]::Show("Le sAMAccountName a été modifié car il existe déjà dans l'AD $($existingUser.DisplayName).`nLe nouveau sAMAccountName est $($samAccountName)", "Attention", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

        # Mise en place de l'email
        $emailuser = TransformUserEmail($emailuser)

        # Création du mot de passe
        $Secure_String_Pwd = ConvertTo-SecureString "InsanePassword" -AsPlainText -Force

        # Vérification si le service est forces de ventes
        if ($department.SelectedItem -eq "FORCES DE VENTES") {
            $global:OU = "FORCE DE VENTE,OU=TEST"
        }

        # Vérification du POBox
        if ($global:POBox -ne $null) {
            $POBox = $global:POBox
        }
        else {
            $POBox = ""
        }

        # Récupération du manager
        $manager = Get-ADUser -Filter "SamAccountName -eq '$($manager.Text)'" -Properties distinguishedName -Server $Script:LogonDC

        # Verification du type de contrat
        if ($checkbox_cdd.Checked) {
            $expirationDate = $datePicker.Value.AddDays(2)
            $descriptionExpirationDate = $expirationDate.AddDays(-1)
            $description = $department.Text + " - " + $job.Text + " - CDD - " + $descriptionExpirationDate.ToString("dd/MM/yyyy")
        }
        else {
            $description = $department.Text + " - " + $job.Text
        }
    
        # Création de l'objet user par rapport au propriétés de l'AD
        $user = @{
            Path              = "OU=" + $global:OU + $ouUser
            Name              = $firstName.Text + " " + $lastName.Text 
            Server            = $Script:LogonDC
            Country           = "FR"
            Company           = $global:societe
            Description       = $description
            Department        = $department.Text
            DisplayName       = $firstName.Text + " " + $lastName.Text 
            EmployeeID        = $employeID.Text
            GivenName         = $firstName.Text
            City              = $global:City
            EmailAddress      = $emailuser
            Manager           = $manager
            PostalCode        = $global:PostalCode
            POBox             = $global:POBox
            SamAccountName    = $samAccountName
            Surname           = $lastName.Text 
            State             = $global:State
            StreetAddress     = $global:StreetAddress
            UserPrincipalName = $emailuser
            Title             = $job.Text
            AccountPassword   = $Secure_String_Pwd
            Enable            = $true
        }

        if ($checkbox_cdd.Checked) {
            $user.Add("AccountExpirationDate", $expirationDate)
        }

        if ($azuread.Checked) {
            try {
                # Création de l'utilisateur dans l'AD
                New-ADUser @user

                # Verification si le service est forces de ventes
                if ($department.SelectedItem -eq "FORCES DE VENTES") {
                    FDVSetPassword $samAccountName
                }

                # Message de confirmation de la création de l'utilisateur dans l'AD
                $result = [System.Windows.Forms.MessageBox]::Show("Utilisateur créé dans l'AD.`nLancement du script de création de boîte mail...", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                if ($result -eq "OK") {

                    Write-Host "`nCréation de la boîte mail en cours..." -ForegroundColor Yellow

                    #Préparation pour le renseignement de la propriété msExchExtensionCustomAttribute1 pour le filtrage AD CONNECT*
                    $userDetails = Get-ADUser -Identity $samAccountName -Properties * -Server $Script:LogonDC

                    $valueADConnectFilter = "O365"
                    $ADConnectFilter = $userDetails
                    $ADConnectFilter."msExchExtensionCustomAttribute1" = $valueADConnectFilter

                    #Récupération des adresses proxy existantes
                    $proxyaddress = Get-ADUser -identity $samAccountName -Properties proxyaddresses -Server $Script:LogonDC | % { $_.proxyaddresses }

                    #Peuplement de la propriété msExchExtensionCustomAttribute1: (filtrage AD Connect)
                    TRY {
                        Set-ADUser -Instance $ADConnectFilter -Server $Script:LogonDC
                        Write-Host "`n`tL'attribut msExchExtensionCustomAttribute1 a bien été renseigné pour l'utilisateur : $($userDetails.DisplayName)" -ForegroundColor Green
                    }
                    CATCH {
                        Write-Host "`n`tL'attribut msExchExtensionCustomAttribute1 n'a pas pu être renseigné pour l'utilisateur : $($userDetails.DisplayName)" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }

                    #Peuplement de la propriété proxyaddress
                    TRY {
                        Set-ADUser -identity $samAccountName -add @{proxyaddresses = "SMTP:$emailuser" } -Server $Script:LogonDC
                        Write-Host "`tSucces du renseignement de l'attribut proxyaddress pour l'utilisateur : $($userDetails.DisplayName)" -ForegroundColor Green
                    }
                    CATCH {
                        Write-Host "`tErreur lors du renseignement de l'attribut proxyaddress pour l'utilisateur $($userDetails.DisplayName)" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }

                    # Lancement d'une replication 
                    Write-Host "`n`tLancement d'une replication Active Directory (domaine 'exemple.test')..." -ForegroundColor Yellow
                    Sync-AllDomainController

                    # Lancement d'une synchronisation AD Connect
                    Write-Host "`n`n`tLancement d'une synchronisation AD Connect..." -ForegroundColor Yellow
                    ADsynchro
    
                    #Temporisation le temps de la replication AD
                    patientez 60

                    #Vérification de l'existance de l'utilisateur
                    TRY {
                        $ADUSERExist = Get-MgUser -UserId $emailuser -ErrorAction SilentlyContinue
                    }
                    CATCH {
                        #Utilisateur Online n'existe pas > upn= $userPrincipalName"
                        $result = [System.Windows.Forms.MessageBox]::Show("L'utilisateur n'a pas été trouvé dans Office 365.`nMerci de vérifier que l'utilisateur a bien été créé dans l'AD.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        
                        if($result -eq "OK") {
                            $form.Close()
                            Disconnect-ExchangeOnline -Confirm:$false
                            Disconnect-MgGraph
                        }
                    }
    
                    # Définition de l'emplacement 
                    Update-MgUser -UserId $emailuser -UserPrincipalName $emailuser -UsageLocation "FR"

                    # Si plus de licence E3 disponible on quitte le script
                    if ($NbE3 -eq 0) {
                        $result = [System.Windows.Forms.MessageBox]::Show("Plus de licence E3 disponible fermeture du script", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        if ($result -eq "OK") {
                            $form.Close()
                            $errorCount ++
                        }
                    } 
                
                    # Sinon on ajoute une licence E3
                    else {
                        Write-Host "`nAjout d'une licence E3 en cours..." -ForegroundColor Yellow
                        TRY {
                            Set-MgUserLicense -UserId $emailuser -AddLicenses $addE3Licenses -RemoveLicenses @()
                            Write-Host "`tLicence E3 ajoutée avec succès" -ForegroundColor Green
                        }
                        CATCH {
                            Write-Host "`n`tErreur lors de l'ajout de la licence E3" -ForegroundColor Red
                        }
                        #si plus de licence P1
                        if ($NbP1 -le "0") {
                            Write-Host "`n`tATTENTION pas de licence Azure AD (P1) disponible" -Foregroundcolor Red
                        }
                        #si P1 dispo, on affecte une P1
                        else {
                            Write-Host "`nAjout d'une licence P1 en cours..." -ForegroundColor Yellow
                            TRY { 
                                Set-MgUserLicense -UserId $emailuser -AddLicenses $addP1Licenses -RemoveLicenses @()
                                Write-Host "`tLicence P1 ajoutée avec succès" -ForegroundColor Green
                            }
                            CATCH {
                                Write-Host "`n`tErreur lors de l'ajout de la licence P1" -ForegroundColor Red
                            } 
                        }
                    }
                
                    Write-Host "`nLa préparation de la boite aux lettres utilisateur en cours" -ForegroundColor Yellow
                    patientez 30
                    Write-Host "`nNous avons bientôt terminé" -ForegroundColor Yellow
                    patientez 30

                    # Initialisation de $existMailbox à $null
                    $existMailbox = $null
                    # Initialisation du compteur de temps d'attente
                    $seconds = 0
                    # Boucle tant que $existMailbox est $null (c'est-à-dire, la boîte aux lettres n'existe pas)
                    while (-not $existMailbox) {
                        $existMailbox = Get-EXOMailbox -Identity $emailuser -ErrorAction SilentlyContinue
                        if (-not $existMailbox) {
                            # Attendre 15 secondes avant de réessayer
                            Write-Host "`nLa boîte aux lettres n'est pas encore remontée, nouvelle tentative dans 15 secondes..."
                            patientez 15
                            # Augmenter le compteur de temps d'attente
                            $seconds += 15
                        }
                    }
                    Write-Host "`nLa boîte aux lettres a été trouvée après $seconds secondes d'attente." -ForegroundColor Green
                    Write-Host "`nFinalisation des propriétés de la boite aux lettres..." -ForegroundColor Yellow
                    Write-host "`tDéfinition de la langue" -ForegroundColor Green
                    Set-Mailbox -Identity $emailuser -Languages "fr-FR"
                    Write-host "`tDéfinition du fuseau horaire" -ForegroundColor Green
                    Set-MailboxRegionalConfiguration -Identity $emailuser -Timezone "Romance Standard Time" -Language "fr-FR" -LocalizeDefaultFolderName:$true
                    Write-host "`tDéfinition des droits calendrier" -ForegroundColor Green
                    Get-EXOMailbox $emailuser | ForEach-Object { Set-MailboxFolderPermission -Identity "$($_.name):\calendrier" -AccessRights LimitedDetails -User default }
                    Write-Host "`nLa boite aux lettres $emailuser a été créée avec succès" -ForegroundColor Green
                    

                    # Vérification si l'entité est la boulangère & co pour envoyé un mail à Jean DUPONT
                    if ($entity.SelectedItem -eq "La Boulangere & Co" -and $azuread.Checked -and $errorCount -eq 0) {
                        $mail = @{
                            From = 'SI@laboulangere-co.fr'
                            To = 'jean.dupont@laboulangere-co.fr'
                            Subject = "Creation de l'adresse mail de $($emailuser)"
                            Body = "Bonjour,`n`nL'adresse mail de l'utilisateur $($emailuser) a ete cree. `n`nCordialement,`nL'equipe Support"
                            SmtpServer = 'SMTP.SERVER'
                            UseSsl = $false
                        }
                        try {
                            Send-MailMessage @mail
                        }
                        catch {
                            Write-Host "Erreur lors de l'envoi du mail" ForegroundColor Red
                            Write-Host "Merci d'envoyé un mail à Jean DUPONT pour lui informer de la création de $($firstName.Text) $($lastName.Text)" ForegroundColor Red
                            Write-Warning "`t Message d'erreur : " 
                            Write-Host $_.Exception.Message
                        }
                    }
                    elseif ($errorCount -ne 0) {
                        Write-Host "`nIl y a eu une ou plusieurs erreurs lors de la création de l'utilisateur" -ForegroundColor Red
                    }

                    # Message de confirmation de la création de la boîte mail ou message d'erreur si la boîte mail n'a pas été créée
                    if ($existMailbox) {
                        $result = [System.Windows.Forms.MessageBox]::Show("Boîte mail créée.`nVoulez-vous créer un autre compte ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    }
                    else {
                        $result = [System.Windows.Forms.MessageBox]::Show("Erreur : la boîte mail n'a pas été créée", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }

                    if ($result -eq "Yes" -or $result -eq "OK") {
                        # Réinitialisation des champs du formulaire
                        Reset-Form
                    }
                    else {
                        $form.Visible = $false
                        Show-Menu
                    }
                }
            }
            catch {
                # Message d'erreur
                $result = [System.Windows.Forms.MessageBox]::Show("Erreur :`n $($_.Exception.Message)", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            if ($result -eq "OK") {
                $form.Visible = $false
                Show-Menu
            }
        }
        else {
            try {
                # Création de l'utilisateur dans l'AD
                New-ADUser @user

                # Verification si le service est forces de ventes
                if ($department.SelectedItem -eq "FORCES DE VENTES") {
                    FDVSetPassword $samAccountName
                }

                # Message de confirmation
                $result = [System.Windows.Forms.MessageBox]::Show("Compte AD crée.`nVoulez-vous créer un autre compte ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

                if ($result -eq "Yes") {
                    # Réinitialisation des champs du formulaire
                    Reset-Form
                }
                else {
                    $form.Visible = $false
                    Show-Menu
                }
            }
            catch {
                # Message d'erreur
                $result = [System.Windows.Forms.MessageBox]::Show("Erreur :`n$($_.Exception.Message)", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                if ($result -eq "OK") {
                    $form.Visible = $false
                    Show-Menu
                }
            }

        }
    })


# --------------------------------------------------------------------------------
# Affichage du formulaire
# --------------------------------------------------------------------------------

$form.Controls.Add($firstName)
$form.Controls.Add($firstName_label)

$form.Controls.Add($lastName)
$form.Controls.Add($lastName_label)

$form.Controls.Add($entity)
$form.Controls.Add($entity_label)

$form.Controls.Add($employeID)
$form.Controls.Add($employeID_label)

$form.Controls.Add($department)
$form.Controls.Add($department_label)

$form.Controls.Add($job)
$form.Controls.Add($job_label)

$form.Controls.Add($manager)
$form.Controls.Add($manager_label)
$form.Controls.Add($button_search)
$form.Controls.Add($displayname_label)

$form.Controls.Add($button_validate)
$form.Controls.Add($button_cancel)

$form.Controls.Add($azureAD)
$form.Controls.Add($checkbox_cdd)
$form.Controls.Add($datePicker)
$form.Controls.Add($datePicker_label)

$form.Controls.Add($E3Label)
$form.Controls.Add($P1Label)

$form.ShowDialog() > $null