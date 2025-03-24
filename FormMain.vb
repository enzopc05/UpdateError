Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data

Public Class FormMain
    Inherits System.Windows.Forms.Form
    
    ' Déclaration des contrôles
    Private txtNumeroCommande As TextBox
    Private txtNumeroLigneMission As TextBox
    Private btnRechercher As Button
    Private dgvResultats As DataGridView
    Private lblStatus As Label

    ' Constructeur
    Public Sub New()
        InitializeComponent()
        
        ' Ajouter menu Configuration
        AjouterMenuConfiguration()
    End Sub
    
    ' Méthode pour ajouter un menu permettant de changer la configuration
    Private Sub AjouterMenuConfiguration()
        ' Création du menu principal
        Dim mainMenu As New MenuStrip()
        mainMenu.Dock = DockStyle.Top
        
        ' Création du menu Fichier
        Dim menuFichier As New ToolStripMenuItem("Fichier")
        
        ' Création de l'élément de menu Configuration
        Dim menuConfiguration As New ToolStripMenuItem("Configuration de la connexion...")
        AddHandler menuConfiguration.Click, AddressOf MenuConfiguration_Click
        
        ' Création de l'élément de menu Quitter
        Dim menuQuitter As New ToolStripMenuItem("Quitter")
        AddHandler menuQuitter.Click, AddressOf MenuQuitter_Click
        
        ' Ajout des éléments au menu Fichier
        menuFichier.DropDownItems.Add(menuConfiguration)
        menuFichier.DropDownItems.Add(New ToolStripSeparator())
        menuFichier.DropDownItems.Add(menuQuitter)
        
        ' Ajout du menu Fichier au menu principal
        mainMenu.Items.Add(menuFichier)
        
        ' Ajout du menu principal au formulaire
        Me.Controls.Add(mainMenu)
        Me.MainMenuStrip = mainMenu
    End Sub
    
    ' Gestionnaire d'événement pour le menu Configuration
    Private Sub MenuConfiguration_Click(sender As Object, e As EventArgs)
        If FormConnexion.ShowConfigurationDialog() Then
            ' Reconnexion réussie, mettre à jour le statut
            lblStatus.Text = "Configuration de la connexion mise à jour."
        End If
    End Sub
    
    ' Gestionnaire d'événement pour le menu Quitter
    Private Sub MenuQuitter_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' Initialisation des composants
    Private Sub InitializeComponent()
        ' Configuration de la fenêtre
        Me.Text = "Vérification des commandes"
        Me.Size = New System.Drawing.Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Création des labels
        Dim lblNumeroCommande As New Label()
        lblNumeroCommande.Text = "Numéro de commande:"
        lblNumeroCommande.Location = New System.Drawing.Point(20, 20)
        lblNumeroCommande.Size = New System.Drawing.Size(150, 20)

        Dim lblNumeroLigneMission As New Label()
        lblNumeroLigneMission.Text = "Numéro de ligne de mission:"
        lblNumeroLigneMission.Location = New System.Drawing.Point(20, 50)
        lblNumeroLigneMission.Size = New System.Drawing.Size(150, 20)

        ' Création des TextBox
        txtNumeroCommande = New TextBox()
        txtNumeroCommande.Location = New System.Drawing.Point(180, 20)
        txtNumeroCommande.Size = New System.Drawing.Size(150, 20)

        txtNumeroLigneMission = New TextBox()
        txtNumeroLigneMission.Location = New System.Drawing.Point(180, 50)
        txtNumeroLigneMission.Size = New System.Drawing.Size(150, 20)

        ' Création du bouton
        btnRechercher = New Button()
        btnRechercher.Text = "Rechercher"
        btnRechercher.Location = New System.Drawing.Point(350, 35)
        btnRechercher.Size = New System.Drawing.Size(100, 30)
        AddHandler btnRechercher.Click, AddressOf BtnRechercher_Click

        ' Création du DataGridView
        dgvResultats = New DataGridView()
        dgvResultats.Location = New System.Drawing.Point(20, 100)
        dgvResultats.Size = New System.Drawing.Size(750, 400)
        dgvResultats.AllowUserToAddRows = False
        dgvResultats.AllowUserToDeleteRows = False
        dgvResultats.ReadOnly = True
        dgvResultats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvResultats.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultats.RowHeadersVisible = False

        ' Label de statut
        lblStatus = New Label()
        lblStatus.Text = "Prêt"
        lblStatus.Location = New System.Drawing.Point(20, 520)
        lblStatus.Size = New System.Drawing.Size(750, 20)

        ' Ajout des contrôles au formulaire
        Me.Controls.Add(lblNumeroCommande)
        Me.Controls.Add(lblNumeroLigneMission)
        Me.Controls.Add(txtNumeroCommande)
        Me.Controls.Add(txtNumeroLigneMission)
        Me.Controls.Add(btnRechercher)
        Me.Controls.Add(dgvResultats)
        Me.Controls.Add(lblStatus)
    End Sub

    ' Gestionnaire d'événement pour le bouton Rechercher
    Private Sub BtnRechercher_Click(sender As Object, e As EventArgs)
        Try
            ' Récupération des valeurs saisies
            Dim numeroCommande As String = txtNumeroCommande.Text.Trim()
            Dim numeroLigneMission As String = txtNumeroLigneMission.Text.Trim()

            ' Validation des entrées
            If String.IsNullOrEmpty(numeroCommande) Then
                lblStatus.Text = "Veuillez saisir un numéro de commande."
                Return
            End If

            ' Mise à jour du statut
            lblStatus.Text = "Recherche en cours..."
            Application.DoEvents()

            ' Appel à la méthode pour récupérer les données
            Dim resultats = DataAccess.GetLignesCommande(numeroCommande, numeroLigneMission)

            ' Affichage des résultats
            dgvResultats.DataSource = resultats

            ' Application de la mise en forme conditionnelle
            FormaterGrille()

            ' Mise à jour du statut
            lblStatus.Text = $"{resultats.Rows.Count} ligne(s) trouvée(s)."
        Catch ex As Exception
            lblStatus.Text = $"Erreur: {ex.Message}"
            MessageBox.Show($"Une erreur est survenue: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Méthode pour formater la grille selon les conditions d'erreur
    Private Sub FormaterGrille()
        Try
            ' Récupérer la source de données
            Dim dataTable As DataTable = TryCast(dgvResultats.DataSource, DataTable)
            If dataTable Is Nothing Then Return
            
            ' Parcourir les lignes et appliquer le formatage conditionnel
            For Each row As DataGridViewRow In dgvResultats.Rows
                ' Ne pas traiter les lignes qui sont nulles ou n'ont pas de DataBoundItem
                If row.DataBoundItem Is Nothing Then Continue For
                
                ' Obtenir le DataRowView correspondant
                Dim dataRowView As DataRowView = TryCast(row.DataBoundItem, DataRowView)
                If dataRowView Is Nothing Then Continue For
                
                ' Vérifier si la ligne est en erreur selon les critères définis
                Dim enErreur As Boolean = DataAccess.EstEnErreur(dataRowView.Row)
                
                ' Mise en forme conditionnelle
                If enErreur Then
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.LightPink
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.Red
                End If
            Next
            
            ' Ajouter dans le statut le compte des lignes en erreur
            Dim lignesEnErreur As Integer = 0
            For Each row As DataRow In dataTable.Rows
                If DataAccess.EstEnErreur(row) Then
                    lignesEnErreur += 1
                End If
            Next
            
            ' Mettre à jour le statut avec le nombre de lignes en erreur
            lblStatus.Text = $"{dataTable.Rows.Count} ligne(s) trouvée(s), dont {lignesEnErreur} en erreur."
            
        Catch ex As Exception
            MessageBox.Show($"Erreur lors du formatage de la grille: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class