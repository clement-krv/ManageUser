<###############################################################################
Script de création de compte AD et boîte mail
###############################################################################>

# --------------------------------------------------------------------------------
# Variable du script
# --------------------------------------------------------------------------------

# Définition du chemin d'accès au fichier jobs.txt
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$jobsFilePath = Join-Path -Path $scriptPath -ChildPath "../jobs.txt"

# --------------------------------------------------------------------------------
# Fonctions 
# --------------------------------------------------------------------------------

function InfoMessage ($password, $emailuser) {
    $info = New-Object System.Windows.Forms.Form
    $info.Text = "La Boulangère - Information"
    $info.Size = New-Object System.Drawing.Size(460, 200)
    $info.StartPosition = "CenterScreen"
    $info.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
    $info.Icon = $icon

    $text = New-Object System.Windows.Forms.Label
    $text.Location = New-Object System.Drawing.Size(20, 10)
    $text.Size = New-Object System.Drawing.Size(500, 20)
    $text.Text = "L'utilisateur a été créé avec succès !"
    $text.Font = New-Object System.Drawing.Font("Tahoma", 12, [System.Drawing.FontStyle]::Bold)

    $password_info = New-Object System.Windows.Forms.Label
    $password_info.Location = New-Object System.Drawing.Size(20, 40)
    $password_info.Size = New-Object System.Drawing.Size(500, 20)
    $password_info.Text = "Voici le mot de passe généré : $password"

    $email_info = New-Object System.Windows.Forms.Label
    $email_info.Location = New-Object System.Drawing.Size(20, 70)
    $email_info.Size = New-Object System.Drawing.Size(500, 20)
    $email_info.Text = "Voici l'adresse email générée : $emailuser"

    $button_copy = New-Object System.Windows.Forms.Button
    $button_copy.Location = New-Object System.Drawing.Size(140, 100)
    $button_copy.Size = New-Object System.Drawing.Size(100, 40)
    $button_copy.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#22c55e")
    $button_copy.Text = "Copier les infos"

    $button_close = New-Object System.Windows.Forms.Button
    $button_close.Location = New-Object System.Drawing.Size(260, 100)
    $button_close.Size = New-Object System.Drawing.Size(100, 40)
    $button_close.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ef4444")
    $button_close.Text = "Fermer"

    $button_copy.Add_Click({
        $both = $password + "`r`n" + $emailuser
        $both | Set-Clipboard
    })

    $button_close.Add_Click({
        $info.Visible = $false
    })

    $info.Controls.Add($text)
    $info.Controls.Add($password_info)
    $info.Controls.Add($email_info)
    $info.Controls.Add($button_copy)
    $info.Controls.Add($button_close)

    $info.ShowDialog() > $null
}

# --------------------------------------------------------------------------------
# Création de la fenêtre
# --------------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form 
$form.Text = "La Boulangère - Création utilisateur AzureAD"
$form.Size = New-Object System.Drawing.Size(520, 350)
$form.StartPosition = "CenterScreen" 
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

$icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
$form.Icon = $icon

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
        "Beaune Brioche" {$global:societe = "Beaune Brioche"; $global:StreetAddress = "Les Cerisières"; $global:State = "Côte-d'Or"; $global:PostalCode = "21209"; $global:City = "Beaune"; Break }
        "La Boulangere & Co" {$global:societe = "La Boulangère & Co"; $global:StreetAddress = "PA La Mongie - 1 rue du Petit Bocage"; $global:State = "Vendée"; $global:PostalCode = "85140"; $global:City = "Essarts-en-Bocage"; Break }
        "Nor'Pain" {$global:societe = "Nor'Pain"; $global:StreetAddress = "Le Vieux Moulin"; $global:State = "Seine-Maritime"; $global:PostalCode = "76890";$global:City = "Val de Saâne"; Break }
        "Ouest Boulangere" {$global:societe = "Ouest Boulangère"; $global:StreetAddress = "ZA La Buzenière - 10, rue Olivier de Serres"; $global:State = "Vendée"; $global:PostalCode = "85503"; $global:City = "Les Herbiers"; Break }
        "Pain Concept" {$global:societe = "Pain Concept"; $global:StreetAddress = "Parc Atlantique"; $global:State = "Vendée"; $global:PostalCode = "85210"; $global:City = "Sainte Hermine"; Break }
        "Panorient" {$global:societe = "Panorient"; $global:StreetAddress = "ZA Eiffel - 36 rue Eiffel"; $global:State = "Seine-et-Marne"; $global:PostalCode = "77220"; $global:City = "Gretz Armainvilliers"; Break }
        "Viennoiserie Ligerienne" {$global:societe = "Viennoiserie Ligérienne"; $global:StreetAddress = "ZI du Gautreau II - 647 rue Antoine Carême"; $global:State = "Vendée"; $global:PostalCode = "85290"; $global:City = "Beaune"; Break }
        "U7" {$global:societe = "U7"; $global:StreetAddress = "ZA La Folie - 38 rue Charles Tellier"; $global:State = "Vendée"; $global:PostalCode = "85036"; $global:City = "La Chaize-Le-Vicomte"; Break }
    } })

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
$manager_label.Size = New-Object System.Drawing.Size(180, 20)
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
    $global:userAD = Get-ADUser -Filter {SamaccountName -eq $manager.Text} -Properties *

    # Vérifiez si l'utilisateur existe
    if ($global:userAD) {
        # Affichez le DisplayName de l'utilisateur
        $displayname_label.Text = $global:userAD.DisplayName
    }
    else {
        # Recherche 2 : on enlève la première lettre et on recherche avec le reste
        $nom = $manager.Text.Substring(1)
        $users = Get-ADUser -Filter {Surname -eq $nom} -Properties *

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
                $displayname_label.Text = $users.DisplayName
                $global:userAD = $users
            }
            # Si aucun utilisateur n'est trouvé
            else {
                $displayname_label.Text = "Utilisateur non trouvé"
            }
        } elseif ($usersCount -eq 1) {
            $displayname_label.Text = $users.DisplayName
            $global:userAD = $users
        } else {
            $displayname_label.Text = "Utilisateur non trouvé"
        }
    }
})

# Label pour afficher le DisplayName du responsable
$displayname_label = New-Object System.Windows.Forms.Label
$displayname_label.Location = New-Object System.Drawing.Size(20, 200)
$displayname_label.Size = New-Object System.Drawing.Size(100, 40)

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

    # Vérification de la saisie du formulaire
    if ($firstName.Text -eq "" -or $lastName.Text -eq "" -or $entity.SelectedItem -eq "" -or $department.Text -eq "" -or $job.Text -eq "" -or $manager.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Veuillez saisir tout les champs", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    #Verification de la saisie du matricule
    if ($employeID.Text -eq ""){
    $result = [System.Windows.Forms.MessageBox]::Show("Le matricule n'est pas renseigné est-ce normal ?`nSi oui alors le matricule prendra comme valeur `"AAAAAAAAAA`".", "Information", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($result -eq "No") {
            return
        } else{
            $employeID.Text = "AAAAAAAAAA"
        }
    }

    # Mise en place du mot de passe aléatoire 5 lettres (alternant consonne et voyelle) + 5 chiffres
    $consonnes = "bcdfghjklmnpqrstvwxz".ToCharArray()
    $voyelles = "aeiouy".ToCharArray()
    $numbers = "0123456789".ToCharArray()

    # Générer 5 lettres en alternant consonne et voyelle
    $randomLetters = -join ((1..5) | ForEach-Object {
        if ($_ % 2 -eq 0) {
            Get-Random -InputObject $consonnes
        } else {
            Get-Random -InputObject $voyelles
        }
    })

    # Assurer que la première lettre est en majuscule
    $randomLetters = $randomLetters.Substring(0,1).ToUpper() + $randomLetters.Substring(1)

    # Générer 5 chiffres aléatoires
    $randomNumbers = -join ((1..5) | ForEach-Object { Get-Random -InputObject $numbers })

    # Combinez les lettres et les chiffres pour créer le mot de passe
    $password = $randomLetters + $randomNumbers

    # Transformation du nom de famille en majuscule
    $lastName.Text = $lastName.Text.ToUpper()

    # Mise en place de l'email
    $emailuser = TransformUserEmail($emailuser)

    # Création de l'objet user par rapport au propriétés de l'Azure
    $user = @{
        AccountEnabled = $true
        City = $global:City
        CompanyName = $global:societe
        Country = "France"
        Department = $department.Text
        DisplayName = $firstName.Text + " " + $lastName.Text
        EmployeeId = $employeID.Text
        GivenName = $firstName.Text
        JobTitle = $job.Text
        UserPrincipalName = $emailuser
        MailNickname = $emailuser.Split('@')[0]
        PasswordProfile = @{
            Password = $password
        }
        PostalCode = $global:PostalCode
        State = $global:State
        StreetAddress = $global:StreetAddress
        Surname = $lastName.Text
        UserType = "Member"
    }

    # Création de l'utilisateur
    Write-Host "Création de l'utilisateur..." -ForegroundColor Yellow
    try {
        $newUser = New-MgUser -BodyParameter $user
        Write-Host "`n`tUtilisateur créé" -ForegroundColor Green
    }
    catch {
        Write-Host "`n`tErreur lors de la création de l'utilisateur"
        Write-Host $_.Exception.Message
    }
    
    # Récupération du manager
    try {
        $manager = Get-MgUser -UserId $global:userAD.EmailAddress -Select Id
        $NewManager = @{
            "@odata.id"="https://graph.microsoft.com/v1.0/users/$($manager.Id)"
        }
        Write-Host "`n`tManager trouvé" -ForegroundColor Green
    }
    catch {
        Write-Host "`n`tErreur lors de la récupération du manager"
        Write-Host $_.Exception.Message
    }

    # Mise en place du manager 
    try {
        Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter $NewManager
        Write-Host "`n`tManager mis en place" -ForegroundColor Green
    }
    catch {
        Write-Host "`n`tErreur lors de la mise en place du manager"
        Write-Host $_.Exception.Message
    }

    #Message d'information sur le mot de passe et l'email
    InfoMessage $password $emailuser


    # Message de confirmation
    $result = [System.Windows.Forms.MessageBox]::Show("Voulez-vous créer un autre compte ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($result -eq "Yes") {
        # Réinitialisation des champs du formulaire
        Reset-AZForm
    }
    else {
        $form.Visible = $false
        Show-Menu
    }
    
    
})

#--------------------------------------------------------------------------------
# Affichage des données du formulaire
#--------------------------------------------------------------------------------

$form.Controls.Add($firstName_label)
$form.Controls.Add($firstName)
$form.Controls.Add($lastName_label)
$form.Controls.Add($lastName)
$form.Controls.Add($entity_label)
$form.Controls.Add($entity)
$form.Controls.Add($employeID_label)
$form.Controls.Add($employeID)
$form.Controls.Add($department_label)
$form.Controls.Add($department)
$form.Controls.Add($job_label)
$form.Controls.Add($job)
$form.Controls.Add($manager_label)
$form.Controls.Add($manager)
$form.Controls.Add($button_search)
$form.Controls.Add($displayname_label)
$form.Controls.Add($button_cancel)
$form.Controls.Add($button_validate)

$form.ShowDialog() > $null