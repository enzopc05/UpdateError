Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.IO

Public Class FormMain
    Inherits System.Windows.Forms.Form
    
    ' Déclaration des contrôles
    Private txtNumeroCommande As TextBox
    Private txtNumeroLigneMission As TextBox
    Private btnRechercher As Button
    Private dgvResultats As DataGridView
    Private lblStatus As Label
    Private pnlHeader As Panel
    Private lblTitle As Label
    Private picLogo As PictureBox

    ' Couleurs de la charte graphique Eurodislog
    Private ReadOnly colorEurodislogBlue As Color = Color.FromArgb(0, 175, 215)  ' Bleu Eurodislog (#00AFD7)
    
    ' Constructeur
    Public Sub New()
        InitializeComponent()
        AjouterMenuConfiguration()
    End Sub
    
    ' Méthode pour ajouter un menu permettant de changer la configuration
    Private Sub AjouterMenuConfiguration()
        ' Création du menu principal
        Dim mainMenu As New MenuStrip()
        mainMenu.Dock = DockStyle.Top
        mainMenu.BackColor = Color.FromArgb(100, 100, 100)  ' Gris foncé
        
        ' Création du menu Fichier
        Dim menuFichier As New ToolStripMenuItem("Fichier")
        menuFichier.ForeColor = Color.White
        
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
        Me.Text = "UpdateError - Eurodislog"
        Me.Size = New System.Drawing.Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Création du panel d'en-tête avec titre
        pnlHeader = New Panel()
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.Height = 80
        pnlHeader.BackColor = colorEurodislogBlue
        
        ' Création du titre
        lblTitle = New Label()
        lblTitle.Text = "Vérification des commandes - UpdateError"
        lblTitle.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(20, 25)
        lblTitle.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        
        ' Création du logo (à droite du titre)
        picLogo = New PictureBox()
        picLogo.Size = New Size(200, 60)
        picLogo.Location = New Point(900, 10)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        picLogo.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        
        ' Charger le logo depuis les ressources ou un fichier
        Try
            ' Essayer de charger depuis un fichier
            If File.Exists(Path.Combine(Application.StartupPath, "logo_eurodislog.jpg")) Then
                picLogo.Image = Image.FromFile(Path.Combine(Application.StartupPath, "logo_eurodislog.jpg"))
            End If
        Catch ex As Exception
            ' En cas d'erreur, ne pas afficher de logo
            Console.WriteLine($"Erreur lors du chargement du logo: {ex.Message}")
        End Try
        
        ' Ajout des contrôles au panel d'en-tête
        pnlHeader.Controls.Add(lblTitle)
        pnlHeader.Controls.Add(picLogo)

        ' Création des labels
        Dim lblNumeroCommande As New Label()
        lblNumeroCommande.Text = "Numéro de commande:"
        lblNumeroCommande.Location = New System.Drawing.Point(20, 100)
        lblNumeroCommande.Size = New System.Drawing.Size(150, 20)
        lblNumeroCommande.Font = New Font("Segoe UI", 10)

        Dim lblNumeroLigneMission As New Label()
        lblNumeroLigneMission.Text = "Numéro de ligne de mission:"
        lblNumeroLigneMission.Location = New System.Drawing.Point(20, 130)
        lblNumeroLigneMission.Size = New System.Drawing.Size(180, 20)
        lblNumeroLigneMission.Font = New Font("Segoe UI", 10)

        ' Création des TextBox
        txtNumeroCommande = New TextBox()
        txtNumeroCommande.Location = New System.Drawing.Point(200, 100)
        txtNumeroCommande.Size = New System.Drawing.Size(200, 25)
        txtNumeroCommande.Font = New Font("Segoe UI", 10)

        txtNumeroLigneMission = New TextBox()
        txtNumeroLigneMission.Location = New System.Drawing.Point(200, 130)
        txtNumeroLigneMission.Size = New System.Drawing.Size(200, 25)
        txtNumeroLigneMission.Font = New Font("Segoe UI", 10)

        ' Création du bouton
        btnRechercher = New Button()
        btnRechercher.Text = "Rechercher"
        btnRechercher.Location = New System.Drawing.Point(420, 115)
        btnRechercher.Size = New System.Drawing.Size(120, 35)
        btnRechercher.BackColor = colorEurodislogBlue
        btnRechercher.ForeColor = Color.White
        btnRechercher.FlatStyle = FlatStyle.Flat
        btnRechercher.FlatAppearance.BorderSize = 0
        btnRechercher.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnRechercher.Cursor = Cursors.Hand
        AddHandler btnRechercher.Click, AddressOf BtnRechercher_Click

        ' Création du DataGridView
        dgvResultats = New DataGridView()
        dgvResultats.Location = New System.Drawing.Point(20, 180)
        dgvResultats.Size = New System.Drawing.Size(1150, 550)
        dgvResultats.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        dgvResultats.AllowUserToAddRows = False
        dgvResultats.AllowUserToDeleteRows = False
        dgvResultats.ReadOnly = True
        dgvResultats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvResultats.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultats.RowHeadersVisible = False
        dgvResultats.BorderStyle = BorderStyle.Fixed3D
        dgvResultats.Font = New Font("Segoe UI", 9)
        dgvResultats.BackgroundColor = Color.White
        dgvResultats.GridColor = Color.LightGray
        dgvResultats.DefaultCellStyle.SelectionBackColor = colorEurodislogBlue
        dgvResultats.DefaultCellStyle.SelectionForeColor = Color.White
        dgvResultats.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240)
        dgvResultats.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        dgvResultats.ColumnHeadersHeight = 30
        dgvResultats.RowTemplate.Height = 25
        dgvResultats.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)

        ' Label de statut
        lblStatus = New Label()
        lblStatus.Text = "Prêt"
        lblStatus.Location = New System.Drawing.Point(20, 740)
        lblStatus.Size = New System.Drawing.Size(1150, 20)
        lblStatus.Font = New Font("Segoe UI", 9)
        lblStatus.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        ' Ajout des contrôles au formulaire
        Me.Controls.Add(pnlHeader)
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