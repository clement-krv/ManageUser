<###############################################################################
Script d'affichage du menu utilisateur et redirection vers les scripts de gestion
###############################################################################>

#--------------------------------------------------------------------------------
# Fonction de redirection vers le script selectionné
#--------------------------------------------------------------------------------

function Move-Script($script, $form) {
    $form.Visible = $false
    . "$PSScriptRoot\$script.ps1"
}

#--------------------------------------------------------------------------------
# Fonction de création du menu de selection
#--------------------------------------------------------------------------------

function Show-Menu {

    # Création de la fenêtre
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "La Boulangère - Menu Utilisateur"
    $form.Size = New-Object System.Drawing.Size(500,300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedSingle'

    # Mise en place du logo
    $icon = New-Object System.Drawing.Icon "$PSScriptRoot/../bco.ico"
    $form.Icon = $icon

    # Création du label de titre
    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(20,20)
    $title.Size = New-Object System.Drawing.Size(500,20)
    $title.Text = "Application de gestion utilisateurs de la Boulangère"
    $title.Font = New-Object System.Drawing.Font("Tahoma", 12, [System.Drawing.FontStyle]::Bold)

    # Création d'un texte explicatif
    $text = New-Object System.Windows.Forms.Label
    $text.Location = New-Object System.Drawing.Point(20,60)
    $text.Size = New-Object System.Drawing.Size(500,20)
    $text.Text = "Veuillez choisir une action à effectuer :"

    # Création du bouton radio pour la création d'un utilisateur AD et O365
    $radioAdAzure = New-Object System.Windows.Forms.RadioButton
    $radioAdAzure.Location = New-Object System.Drawing.Point(20,100)
    $radioAdAzure.Size = New-Object System.Drawing.Size(500,20)
    $radioAdAzure.Text = "Créer un utilisateur AD et Azure"

    # Création du bouton radio pour la création d'un utilisateur O365
    $radioAzure = New-Object System.Windows.Forms.RadioButton
    $radioAzure.Location = New-Object System.Drawing.Point(20,140)
    $radioAzure.Size = New-Object System.Drawing.Size(500,20)
    $radioAzure.Text = "Créer un utilisateur Azure"

    # Création du bouton radio pour la suppression d'un utilisateur AD et O365
    $radioAdAzureDel = New-Object System.Windows.Forms.RadioButton
    $radioAdAzureDel.Location = New-Object System.Drawing.Point(20,180)
    $radioAdAzureDel.Size = New-Object System.Drawing.Size(500,20)
    $radioAdAzureDel.Text = "Supprimer un utilisateur AD et Azure"

    #Création du bouton d'anulation
    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Location = New-Object System.Drawing.Point(240,210)
    $buttonCancel.Size = New-Object System.Drawing.Size(100,40)
    $buttonCancel.BackColor = "#ef4444"
    $buttonCancel.Text = "Annuler"
    $buttonCancel.Add_Click({
        $form.Close()
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-MgGraph
    })

    # Création du bouton de validation
    $buttonValidate = New-Object System.Windows.Forms.Button
    $buttonValidate.Location = New-Object System.Drawing.Point(360,210)
    $buttonValidate.Size = New-Object System.Drawing.Size(100,40)
    $buttonValidate.BackColor = "#22c55e"
    $buttonValidate.Text = "Valider"

    # Action du bouton de validation
    $buttonValidate.Add_Click({    
        switch ($true) {
            $radioAdAzure.Checked {
                Move-Script "createAdAzure" $form
            }
            $radioAzure.Checked {
                Move-Script "createAzure" $form
            }
            $radioAdAzureDel.Checked {
                Move-Script "deleteAdAzure" $form
            }
        }
    })

    # Ajout des éléments à la fenêtre
    $form.Controls.Add($title)
    $form.Controls.Add($text)
    $form.Controls.Add($radioAdAzure)
    $form.Controls.Add($radioAzure)
    $form.Controls.Add($radioAdAzureDel)
    $form.Controls.Add($buttonCancel)
    $form.Controls.Add($buttonValidate)

    $form.ShowDialog() > $null
}