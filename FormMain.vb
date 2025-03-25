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
    Private pnlSearch As Panel
    Private pnlFooter As Panel

    ' Couleurs de la charte graphique
    Private ReadOnly colorBlue As Color = Color.FromArgb(0, 175, 215)  ' Bleu (#00AFD7)
    Private ReadOnly colorDarkBlue As Color = Color.FromArgb(47, 54, 127)  ' Bleu foncé (#2F367F)
    Private ReadOnly colorLightGray As Color = Color.FromArgb(240, 240, 240)  ' Gris clair pour l'arrière-plan
    
    ' Constructeur
    Public Sub New()
        InitializeComponent()
        AjouterMenuConfiguration()
        
        ' Personnalisation supplémentaire après initialisation
        ConfigurerStyleDataGridView()
        ConfigurerDataGridViewInteractions()
        SetPlaceholders()
    End Sub
    
    ' Méthode pour ajouter des placeholders aux TextBox
    Private Sub SetPlaceholders()
        AddPlaceholder(txtNumeroCommande, "Saisissez le numéro de commande...")
        AddPlaceholder(txtNumeroLigneMission, "Optionnel - Numéro de ligne...")
    End Sub
    
    ' Méthode générique pour ajouter un placeholder à un TextBox
    Private Sub AddPlaceholder(textBox As TextBox, placeholderText As String)
        textBox.Tag = placeholderText
        textBox.Text = placeholderText
        textBox.ForeColor = Color.Gray
        
        AddHandler textBox.Enter, Sub(sender As Object, e As EventArgs)
            If textBox.Text = textBox.Tag.ToString() Then
                textBox.Text = ""
                textBox.ForeColor = Color.Black
            End If
        End Sub
        
        AddHandler textBox.Leave, Sub(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(textBox.Text) Then
                textBox.Text = textBox.Tag.ToString()
                textBox.ForeColor = Color.Gray
            End If
        End Sub
    End Sub
    
    ' Méthode pour ajouter un menu permettant de changer la configuration
    Private Sub AjouterMenuConfiguration()
        ' Création du menu principal
        Dim mainMenu As New MenuStrip()
        mainMenu.Dock = DockStyle.Top
        mainMenu.BackColor = colorDarkBlue  ' Bleu foncé pour le menu
        mainMenu.ForeColor = Color.White
        mainMenu.Padding = New Padding(10, 2, 0, 2)
        
        ' Création du menu Fichier
        Dim menuFichier As New ToolStripMenuItem("Fichier")
        menuFichier.ForeColor = Color.White
        menuFichier.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        
        ' Création de l'élément de menu Configuration
        Dim menuConfiguration As New ToolStripMenuItem("Configuration de la connexion...")
        menuConfiguration.Image = GetIcon("settings")
        AddHandler menuConfiguration.Click, AddressOf MenuConfiguration_Click
        
        ' Création de l'élément de menu Quitter
        Dim menuQuitter As New ToolStripMenuItem("Quitter")
        menuQuitter.Image = GetIcon("exit")
        AddHandler menuQuitter.Click, AddressOf MenuQuitter_Click
        
        ' Ajout des éléments au menu Fichier
        menuFichier.DropDownItems.Add(menuConfiguration)
        menuFichier.DropDownItems.Add(New ToolStripSeparator())
        menuFichier.DropDownItems.Add(menuQuitter)
        
        ' Création du menu Aide
        Dim menuAide As New ToolStripMenuItem("Aide")
        menuAide.ForeColor = Color.White
        menuAide.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        
        ' Création de l'élément de menu À propos
        Dim menuAPropos As New ToolStripMenuItem("À propos...")
        menuAPropos.Image = GetIcon("info")
        AddHandler menuAPropos.Click, AddressOf MenuAPropos_Click
        
        ' Ajout des éléments au menu Aide
        menuAide.DropDownItems.Add(menuAPropos)
        
        ' Ajout des menus au menu principal
        mainMenu.Items.Add(menuFichier)
        mainMenu.Items.Add(menuAide)
        
        ' Ajout du menu principal au formulaire
        Me.Controls.Add(mainMenu)
        Me.MainMenuStrip = mainMenu
    End Sub
    
    ' Fonction pour obtenir une icône pour le menu (simulée pour l'exemple)
    Private Function GetIcon(iconName As String) As Image
        ' Dans une application réelle, vous chargeriez les icônes depuis les ressources
        ' Ici, on crée juste des icônes de remplacement
        Dim icon As New Bitmap(16, 16)
        Using g As Graphics = Graphics.FromImage(icon)
            g.Clear(Color.Transparent)
            g.DrawRectangle(New Pen(Color.White), 2, 2, 12, 12)
        End Using
        Return icon
    End Function
    
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
    
    ' Gestionnaire d'événement pour le menu À propos
    Private Sub MenuAPropos_Click(sender As Object, e As EventArgs)
        MessageBox.Show("UpdateError - Version 1.0.0" & Environment.NewLine & 
                     "© 2025 Eurodislog" & Environment.NewLine & 
                     "Application de vérification des commandes", 
                     "À propos de UpdateError", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Initialisation des composants
    Private Sub InitializeComponent()
        ' Configuration de la fenêtre
        Me.Text = "UpdateError - Eurodislog"
        Me.Size = New System.Drawing.Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = colorLightGray  ' Fond gris clair pour l'application
        Me.Font = New Font("Segoe UI", 9)

        ' Création du panel d'en-tête avec titre et logo
        pnlHeader = New Panel()
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.Height = 80
        pnlHeader.BackColor = colorBlue
        pnlHeader.Padding = New Padding(15, 5, 15, 5)
        
        ' Création du titre
        lblTitle = New Label()
        lblTitle.Text = "Vérification des commandes - UpdateError"
        lblTitle.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(20, 25)
        lblTitle.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom
        
        ' Création du logo (à droite du titre)
        picLogo = New PictureBox()
        picLogo.Size = New Size(200, 60)
        picLogo.Location = New Point(pnlHeader.Width - 220, 10)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        picLogo.Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        
        ' Charger le logo depuis les ressources ou un fichier
        Try
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
        
        ' Création du panel de recherche
        pnlSearch = New Panel()
        pnlSearch.Dock = DockStyle.Top
        pnlSearch.Height = 100
        pnlSearch.BackColor = Color.White
        pnlSearch.Padding = New Padding(15, 15, 15, 15)
        pnlSearch.BorderStyle = BorderStyle.None
        pnlSearch.Margin = New Padding(0, 20, 0, 20)
        
        ' Création de l'ombre pour le panel
        AddHandler pnlSearch.Paint, Sub(sender As Object, e As PaintEventArgs)
            Dim shadowRect As New Rectangle(0, pnlSearch.Height - 5, pnlSearch.Width, 5)
            Using brush As New Drawing2D.LinearGradientBrush(
                shadowRect, 
                Color.FromArgb(20, 0, 0, 0), 
                Color.FromArgb(0, 0, 0, 0), 
                Drawing2D.LinearGradientMode.Vertical)
                e.Graphics.FillRectangle(brush, shadowRect)
            End Using
        End Sub

        ' Création des labels avec texte complet et taille adéquate
        Dim lblNumeroCommande As New Label()
        lblNumeroCommande.Text = "Numéro de commande"
        lblNumeroCommande.Location = New System.Drawing.Point(20, 20)
        lblNumeroCommande.Size = New System.Drawing.Size(250, 20)  ' Augmentation de la largeur
        lblNumeroCommande.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblNumeroCommande.ForeColor = colorDarkBlue
        lblNumeroCommande.AutoSize = True  ' Permettre au texte de s'afficher en entier

        Dim lblNumeroLigneMission As New Label()
        lblNumeroLigneMission.Text = "Numéro de ligne de mission"
        lblNumeroLigneMission.Location = New System.Drawing.Point(350, 20)
        lblNumeroLigneMission.Size = New System.Drawing.Size(250, 20)  ' Taille cohérente
        lblNumeroLigneMission.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        lblNumeroLigneMission.ForeColor = colorDarkBlue
        lblNumeroLigneMission.AutoSize = True  ' Permettre au texte de s'afficher en entier

        ' Création des TextBox avec style amélioré et espacement correct
        txtNumeroCommande = New TextBox()
        txtNumeroCommande.Location = New System.Drawing.Point(20, 45)
        txtNumeroCommande.Size = New System.Drawing.Size(280, 30)  ' Augmentation de la largeur
        txtNumeroCommande.Font = New Font("Segoe UI", 11)
        txtNumeroCommande.BorderStyle = BorderStyle.FixedSingle
        txtNumeroCommande.Padding = New Padding(5)

        txtNumeroLigneMission = New TextBox()
        txtNumeroLigneMission.Location = New System.Drawing.Point(350, 45)
        txtNumeroLigneMission.Size = New System.Drawing.Size(280, 30)  ' Cohérence avec la première TextBox
        txtNumeroLigneMission.Font = New Font("Segoe UI", 11)
        txtNumeroLigneMission.BorderStyle = BorderStyle.FixedSingle
        txtNumeroLigneMission.Padding = New Padding(5)

        ' Création du bouton avec style amélioré
        btnRechercher = New Button()
        btnRechercher.Text = "Rechercher"
        btnRechercher.Location = New System.Drawing.Point(650, 40)
        btnRechercher.Size = New System.Drawing.Size(140, 40)
        btnRechercher.BackColor = colorBlue
        btnRechercher.ForeColor = Color.White
        btnRechercher.FlatStyle = FlatStyle.Flat
        btnRechercher.FlatAppearance.BorderSize = 0
        btnRechercher.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        btnRechercher.Cursor = Cursors.Hand
        btnRechercher.Image = GetIcon("search")
        btnRechercher.ImageAlign = ContentAlignment.MiddleLeft
        btnRechercher.TextAlign = ContentAlignment.MiddleCenter
        btnRechercher.TextImageRelation = TextImageRelation.ImageBeforeText
        btnRechercher.Padding = New Padding(10, 0, 10, 0)
        
        ' Effet de survol pour le bouton
        AddHandler btnRechercher.MouseEnter, Sub(sender As Object, e As EventArgs)
            btnRechercher.BackColor = colorDarkBlue
        End Sub
        
        AddHandler btnRechercher.MouseLeave, Sub(sender As Object, e As EventArgs)
            btnRechercher.BackColor = colorBlue
        End Sub
        
        AddHandler btnRechercher.Click, AddressOf BtnRechercher_Click

        ' Ajout des contrôles au panel de recherche
        pnlSearch.Controls.Add(lblNumeroCommande)
        pnlSearch.Controls.Add(lblNumeroLigneMission)
        pnlSearch.Controls.Add(txtNumeroCommande)
        pnlSearch.Controls.Add(txtNumeroLigneMission)
        pnlSearch.Controls.Add(btnRechercher)

        ' Création du DataGridView avec style amélioré
        dgvResultats = New DataGridView()
        dgvResultats.Location = New System.Drawing.Point(20, 200)
        dgvResultats.Size = New System.Drawing.Size(1150, 500)
        dgvResultats.Dock = DockStyle.Fill
        dgvResultats.Margin = New Padding(20)
        dgvResultats.Padding = New Padding(10)
        
        ' Création du panel de pied de page pour le statut
        pnlFooter = New Panel()
        pnlFooter.Dock = DockStyle.Bottom
        pnlFooter.Height = 40
        pnlFooter.BackColor = Color.White
        pnlFooter.BorderStyle = BorderStyle.None
        
        ' Ajouter une ombre au pied de page
        AddHandler pnlFooter.Paint, Sub(sender As Object, e As PaintEventArgs)
            Dim shadowRect As New Rectangle(0, 0, pnlFooter.Width, 5)
            Using brush As New Drawing2D.LinearGradientBrush(
                shadowRect, 
                Color.FromArgb(0, 0, 0, 0), 
                Color.FromArgb(20, 0, 0, 0), 
                Drawing2D.LinearGradientMode.Vertical)
                e.Graphics.FillRectangle(brush, shadowRect)
            End Using
        End Sub

        ' Label de statut
        lblStatus = New Label()
        lblStatus.Text = "Prêt"
        lblStatus.Dock = DockStyle.Fill
        lblStatus.TextAlign = ContentAlignment.MiddleLeft
        lblStatus.Font = New Font("Segoe UI", 9)
        lblStatus.Padding = New Padding(20, 0, 0, 0)
        
        ' Ajout du label de statut au panel de pied de page
        pnlFooter.Controls.Add(lblStatus)

        ' Création d'un panel principal pour contenir le DataGridView avec marge
        Dim pnlMain As New Panel()
        pnlMain.Dock = DockStyle.Fill
        pnlMain.Padding = New Padding(20)
        pnlMain.Controls.Add(dgvResultats)

        ' Ajout des contrôles au formulaire dans l'ordre correct
        Me.Controls.Add(pnlFooter)
        Me.Controls.Add(pnlMain)
        Me.Controls.Add(pnlSearch)
        Me.Controls.Add(pnlHeader)
    End Sub
    
    ' Méthode pour configurer le style du DataGridView
    Private Sub ConfigurerStyleDataGridView()
        dgvResultats.AllowUserToAddRows = False
        dgvResultats.AllowUserToDeleteRows = False
        dgvResultats.ReadOnly = True
        dgvResultats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvResultats.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultats.RowHeadersVisible = False
        dgvResultats.BorderStyle = BorderStyle.None
        dgvResultats.BackgroundColor = Color.White
        dgvResultats.GridColor = Color.FromArgb(230, 230, 230)
        dgvResultats.DefaultCellStyle.SelectionBackColor = colorBlue
        dgvResultats.DefaultCellStyle.SelectionForeColor = Color.White
        dgvResultats.ColumnHeadersDefaultCellStyle.BackColor = colorDarkBlue
        dgvResultats.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvResultats.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        dgvResultats.ColumnHeadersHeight = 40
        dgvResultats.RowTemplate.Height = 30
        dgvResultats.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgvResultats.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgvResultats.EnableHeadersVisualStyles = False
    End Sub
    
    ' Méthode pour configurer le DataGridView avec un menu contextuel
    Private Sub ConfigurerDataGridViewInteractions()
        ' Permettre la sélection de lignes
        dgvResultats.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultats.MultiSelect = False
        
        ' Ajouter un gestionnaire pour le double-clic
        AddHandler dgvResultats.CellDoubleClick, AddressOf DgvResultats_CellDoubleClick
        
        ' Ajouter un menu contextuel
        Dim contextMenu As New ContextMenuStrip()
        
        ' Option pour éditer le numéro de tracking
        Dim menuEditTracking As New ToolStripMenuItem("Modifier le numéro de tracking")
        menuEditTracking.Image = GetIcon("edit")
        AddHandler menuEditTracking.Click, AddressOf MenuEditTracking_Click
        contextMenu.Items.Add(menuEditTracking)
        
        ' Ajouter le menu contextuel au DataGridView
        dgvResultats.ContextMenuStrip = contextMenu
        
        ' Ajouter un gestionnaire pour afficher le menu contextuel uniquement sur les lignes valides
        AddHandler dgvResultats.CellContextMenuStripNeeded, Sub(sender As Object, e As DataGridViewCellContextMenuStripNeededEventArgs)
            If e.RowIndex >= 0 Then
                e.ContextMenuStrip = contextMenu
            Else
                e.ContextMenuStrip = Nothing
            End If
        End Sub
    End Sub
    
    ' Gestionnaire d'événement pour le double-clic sur une cellule
    Private Sub DgvResultats_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        ' Vérifier que la cellule est valide
        If e.RowIndex < 0 Then Return
        
        ' Si la colonne cliquée est EXO_TRAK, ouvrir le formulaire d'édition
        If dgvResultats.Columns(e.ColumnIndex).Name = "EXO_TRAK" Then
            EditerExoTrakPourLigneSelectionnee()
        End If
    End Sub
    
    ' Gestionnaire d'événement pour l'option du menu contextuel
    Private Sub MenuEditTracking_Click(sender As Object, e As EventArgs)
        EditerExoTrakPourLigneSelectionnee()
    End Sub
    
    ' Méthode pour éditer EXO_TRAK pour la ligne sélectionnée
    Private Sub EditerExoTrakPourLigneSelectionnee()
        Try
            ' Vérifier qu'une ligne est sélectionnée
            If dgvResultats.SelectedRows.Count = 0 Then
                MessageBox.Show("Veuillez sélectionner une ligne à modifier.", 
                              "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            
            ' Récupérer les données de la ligne sélectionnée
            Dim row As DataGridViewRow = dgvResultats.SelectedRows(0)
            
            ' Obtenir les valeurs nécessaires
            Dim exoKeyU As String = row.Cells("EXO_KEYU").Value.ToString()
Dim exoTrak As String = ""
            
' Vérifier si la valeur est null avant de la convertir en string
If row.Cells("EXO_TRAK").Value IsNot Nothing Then
    exoTrak = row.Cells("EXO_TRAK").Value.ToString()
End If            
            ' Ouvrir le formulaire d'édition
            Dim result = FormEditionExoTrak.EditExoTrak(exoKeyU, exoTrak)
            
            ' Si la modification a été confirmée, mettre à jour l'affichage
            If result.Item1 Then
                ' Mettre à jour la valeur dans la grille
                row.Cells("EXO_TRAK").Value = result.Item2
                
                ' Rafraîchir la mise en forme conditionnelle
                FormaterGrille()
                
                ' Afficher un message de confirmation
                lblStatus.Text = $"Le numéro de tracking pour la ligne {exoKeyU} a été mis à jour."
                lblStatus.ForeColor = Color.FromArgb(0, 150, 0)  ' Vert
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erreur lors de l'édition du numéro de tracking: {ex.Message}", 
                          "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Gestionnaire d'événement pour le bouton Rechercher
    Private Sub BtnRechercher_Click(sender As Object, e As EventArgs)
        Try
            ' Récupération des valeurs saisies
            Dim numeroCommande As String = txtNumeroCommande.Text.Trim()
            Dim numeroLigneMission As String = txtNumeroLigneMission.Text.Trim()
            
            ' Ignorer les placeholders
            If numeroCommande = txtNumeroCommande.Tag.ToString() Then
                numeroCommande = ""
            End If
            
            If numeroLigneMission = txtNumeroLigneMission.Tag.ToString() Then
                numeroLigneMission = ""
            End If

            ' Validation des entrées
            If String.IsNullOrEmpty(numeroCommande) Then
                lblStatus.Text = "Veuillez saisir un numéro de commande."
                txtNumeroCommande.Focus()
                return
            End If

            ' Mise à jour du statut avec indicateur visuel
            lblStatus.Text = "Recherche en cours..."
            lblStatus.ForeColor = colorBlue
            Application.DoEvents()

            ' Appel à la méthode pour récupérer les données
            Dim resultats = DataAccess.GetLignesCommande(numeroCommande, numeroLigneMission)

            ' Affichage des résultats
            dgvResultats.DataSource = resultats

            ' Application de la mise en forme conditionnelle
            FormaterGrille()

            ' Mise à jour du statut
            Dim message As String = $"{resultats.Rows.Count} ligne(s) trouvée(s)."
            lblStatus.Text = message
            lblStatus.ForeColor = Color.Black
            
            ' Permettre l'édition si des lignes ont été trouvées
            If resultats.Rows.Count > 0 Then
                dgvResultats.ClearSelection()  ' Enlever la sélection par défaut
                lblStatus.Text += " Double-cliquez sur un numéro de tracking pour le modifier."
            End If
        Catch ex As Exception
            lblStatus.Text = $"Erreur: {ex.Message}"
            lblStatus.ForeColor = Color.Red
            MessageBox.Show($"Une erreur est survenue: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Méthode pour formater la grille selon les conditions d'erreur
    Private Sub FormaterGrille()
        Try
            ' Récupérer la source de données
            Dim dataTable As DataTable = TryCast(dgvResultats.DataSource, DataTable)
            If dataTable Is Nothing Then Return
            
            ' Configurer les noms des colonnes et leur visibilité
            ConfigurerColonnesDataGridView()
            
            ' Parcourir les lignes et appliquer le formatage conditionnel
            For Each row As DataGridViewRow In dgvResultats.Rows
                ' Ne pas traiter les lignes qui sont nulles ou n'ont pas de DataBoundItem
                If row.DataBoundItem Is Nothing Then Continue For
                
                ' Obtenir le DataRowView correspondant
                Dim dataRowView As DataRowView = TryCast(row.DataBoundItem, DataRowView)
                If dataRowView Is Nothing Then Continue For
                
                ' Vérifier si la ligne est en erreur selon les critères définis
                Dim enErreur As Boolean = DataAccess.EstEnErreur(dataRowView.Row)
                
                ' Mise en forme conditionnelle avec style amélioré
                If enErreur Then
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 238)  ' Rose pâle
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(213, 0, 0)  ' Rouge foncé
                    row.DefaultCellStyle.Font = New Font(dgvResultats.DefaultCellStyle.Font, FontStyle.Bold)
                    
                    ' Ajouter une info-bulle pour indiquer la raison de l'erreur
                    For Each cell As DataGridViewCell In row.Cells
                        ' Ajouter l'info-bulle uniquement si la cellule EXO_TRAK est concernée
                        If cell.OwningColumn.Name = "EXO_TRAK" Then
                            cell.ToolTipText = "La longueur du code de suivi doit être de 9 ou 16 caractères."
                        End If
                    Next
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
            
            ' Colorier le statut en fonction des erreurs
            If lignesEnErreur > 0 Then
                lblStatus.ForeColor = Color.FromArgb(213, 0, 0)  ' Rouge
            Else
                lblStatus.ForeColor = Color.FromArgb(0, 150, 0)  ' Vert
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erreur lors du formatage de la grille: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Méthode pour configurer les colonnes du DataGridView
    Private Sub ConfigurerColonnesDataGridView()
    Private Sub ConfigurerColonnesDataGridView()
        Try
            ' Vérifier que le DataGridView a une source de données
            If dgvResultats.DataSource Is Nothing Then Return
            
            ' Définir les titres des colonnes et leur ordre
            If dgvResultats.Columns.Contains("EXO_KEYU") Then
                dgvResultats.Columns("EXO_KEYU").HeaderText = "N° Ligne"
                dgvResultats.Columns("EXO_KEYU").DisplayIndex = 0
                dgvResultats.Columns("EXO_KEYU").Width = 100
            End If
            
            If dgvResultats.Columns.Contains("OPE_NOOE") Then
                dgvResultats.Columns("OPE_NOOE").HeaderText = "N° Commande"
                dgvResultats.Columns("OPE_NOOE").DisplayIndex = 1
                dgvResultats.Columns("OPE_NOOE").Width = 120
            End If
            
            If dgvResultats.Columns.Contains("MIL_NOLM") Then
                dgvResultats.Columns("MIL_NOLM").HeaderText = "N° Ligne Mission"
                dgvResultats.Columns("MIL_NOLM").DisplayIndex = 2
                dgvResultats.Columns("MIL_NOLM").Width = 120
            End If
            
            If dgvResultats.Columns.Contains("EXO_SCAN") Then
                dgvResultats.Columns("EXO_SCAN").HeaderText = "Code Article"
                dgvResultats.Columns("EXO_SCAN").DisplayIndex = 3
                dgvResultats.Columns("EXO_SCAN").Width = 150
            End If
            
            If dgvResultats.Columns.Contains("EXO_TRAK") Then
                dgvResultats.Columns("EXO_TRAK").HeaderText = "N° Traçabilité"
                dgvResultats.Columns("EXO_TRAK").DisplayIndex = 4
                dgvResultats.Columns("EXO_TRAK").Width = 150
                
                ' Ajouter une mise en forme conditionnelle à la colonne de tracking
                Dim column As DataGridViewColumn = dgvResultats.Columns("EXO_TRAK")
                column.DefaultCellStyle.Padding = New Padding(5, 0, 5, 0)
            End If
            
            ' Optimiser l'affichage des colonnes
            dgvResultats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            
        Catch ex As Exception
            MessageBox.Show($"Erreur lors de la configuration des colonnes: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class