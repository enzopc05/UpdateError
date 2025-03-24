Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

Public Class FormConnexion
    Inherits System.Windows.Forms.Form
    
    ' Déclaration des contrôles
    Private txtServeur As TextBox
    Private txtBaseDeDonnees As TextBox
    Private chkAuthentificationWindows As CheckBox
    Private txtUtilisateur As TextBox
    Private txtMotDePasse As TextBox
    Private btnTester As Button
    Private btnEnregistrer As Button
    Private btnAnnuler As Button
    Private lblStatus As Label
    
    ' Fichier de configuration pour stocker les paramètres
    Private Shared ReadOnly ConfigFile As String = Path.Combine(Application.StartupPath, "config.ini")
    
    ' Constructeur
    Public Sub New()
        InitializeComponent()
        ChargerConfiguration()
    End Sub
    
    ' Initialisation des composants
    Private Sub InitializeComponent()
        ' Configuration de la fenêtre
        Me.Text = "Configuration de la connexion à la base de données"
        Me.Size = New System.Drawing.Size(500, 350)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        
        ' Création des labels
        Dim lblServeur As New Label()
        lblServeur.Text = "Serveur SQL Server:"
        lblServeur.Location = New System.Drawing.Point(20, 20)
        lblServeur.Size = New System.Drawing.Size(150, 20)
        
        Dim lblBaseDeDonnees As New Label()
        lblBaseDeDonnees.Text = "Base de données:"
        lblBaseDeDonnees.Location = New System.Drawing.Point(20, 50)
        lblBaseDeDonnees.Size = New System.Drawing.Size(150, 20)
        
        Dim lblAuthentification As New Label()
        lblAuthentification.Text = "Authentification:"
        lblAuthentification.Location = New System.Drawing.Point(20, 80)
        lblAuthentification.Size = New System.Drawing.Size(150, 20)
        
        Dim lblUtilisateur As New Label()
        lblUtilisateur.Text = "Utilisateur SQL:"
        lblUtilisateur.Location = New System.Drawing.Point(20, 110)
        lblUtilisateur.Size = New System.Drawing.Size(150, 20)
        
        Dim lblMotDePasse As New Label()
        lblMotDePasse.Text = "Mot de passe SQL:"
        lblMotDePasse.Location = New System.Drawing.Point(20, 140)
        lblMotDePasse.Size = New System.Drawing.Size(150, 20)
        
        ' Création des TextBox
        txtServeur = New TextBox()
        txtServeur.Location = New System.Drawing.Point(180, 20)
        txtServeur.Size = New System.Drawing.Size(280, 20)
        
        txtBaseDeDonnees = New TextBox()
        txtBaseDeDonnees.Location = New System.Drawing.Point(180, 50)
        txtBaseDeDonnees.Size = New System.Drawing.Size(280, 20)
        
        chkAuthentificationWindows = New CheckBox()
        chkAuthentificationWindows.Text = "Utiliser l'authentification Windows"
        chkAuthentificationWindows.Location = New System.Drawing.Point(180, 80)
        chkAuthentificationWindows.Size = New System.Drawing.Size(280, 20)
        chkAuthentificationWindows.Checked = True
        AddHandler chkAuthentificationWindows.CheckedChanged, AddressOf ChkAuthentificationWindows_CheckedChanged
        
        txtUtilisateur = New TextBox()
        txtUtilisateur.Location = New System.Drawing.Point(180, 110)
        txtUtilisateur.Size = New System.Drawing.Size(280, 20)
        txtUtilisateur.Enabled = Not chkAuthentificationWindows.Checked
        
        txtMotDePasse = New TextBox()
        txtMotDePasse.Location = New System.Drawing.Point(180, 140)
        txtMotDePasse.Size = New System.Drawing.Size(280, 20)
        txtMotDePasse.PasswordChar = "*"c
        txtMotDePasse.Enabled = Not chkAuthentificationWindows.Checked
        
        ' Création des boutons
        btnTester = New Button()
        btnTester.Text = "Tester la connexion"
        btnTester.Location = New System.Drawing.Point(180, 180)
        btnTester.Size = New System.Drawing.Size(150, 30)
        AddHandler btnTester.Click, AddressOf BtnTester_Click
        
        btnEnregistrer = New Button()
        btnEnregistrer.Text = "Enregistrer"
        btnEnregistrer.Location = New System.Drawing.Point(100, 240)
        btnEnregistrer.Size = New System.Drawing.Size(100, 30)
        AddHandler btnEnregistrer.Click, AddressOf BtnEnregistrer_Click
        
        btnAnnuler = New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Location = New System.Drawing.Point(280, 240)
        btnAnnuler.Size = New System.Drawing.Size(100, 30)
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        
        ' Label de statut
        lblStatus = New Label()
        lblStatus.Text = "Veuillez configurer la connexion à la base de données."
        lblStatus.Location = New System.Drawing.Point(20, 290)
        lblStatus.Size = New System.Drawing.Size(450, 20)
        
        ' Ajout des contrôles au formulaire
        Me.Controls.Add(lblServeur)
        Me.Controls.Add(lblBaseDeDonnees)
        Me.Controls.Add(lblAuthentification)
        Me.Controls.Add(lblUtilisateur)
        Me.Controls.Add(lblMotDePasse)
        Me.Controls.Add(txtServeur)
        Me.Controls.Add(txtBaseDeDonnees)
        Me.Controls.Add(chkAuthentificationWindows)
        Me.Controls.Add(txtUtilisateur)
        Me.Controls.Add(txtMotDePasse)
        Me.Controls.Add(btnTester)
        Me.Controls.Add(btnEnregistrer)
        Me.Controls.Add(btnAnnuler)
        Me.Controls.Add(lblStatus)
    End Sub
    
    ' Gestionnaire d'événement pour le changement d'état de la checkbox
    Private Sub ChkAuthentificationWindows_CheckedChanged(sender As Object, e As EventArgs)
        txtUtilisateur.Enabled = Not chkAuthentificationWindows.Checked
        txtMotDePasse.Enabled = Not chkAuthentificationWindows.Checked
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Tester
    Private Sub BtnTester_Click(sender As Object, e As EventArgs)
        Try
            ' Récupération des valeurs saisies
            Dim serveur As String = txtServeur.Text.Trim()
            Dim baseDeDonnees As String = txtBaseDeDonnees.Text.Trim()
            Dim authentificationWindows As Boolean = chkAuthentificationWindows.Checked
            Dim utilisateur As String = txtUtilisateur.Text.Trim()
            Dim motDePasse As String = txtMotDePasse.Text
            
            ' Validation basique des entrées
            If String.IsNullOrEmpty(serveur) Then
                lblStatus.Text = "Veuillez saisir le nom du serveur."
                Return
            End If
            
            If String.IsNullOrEmpty(baseDeDonnees) Then
                lblStatus.Text = "Veuillez saisir le nom de la base de données."
                Return
            End If
            
            If Not authentificationWindows AndAlso String.IsNullOrEmpty(utilisateur) Then
                lblStatus.Text = "Veuillez saisir un nom d'utilisateur SQL."
                Return
            End If
            
            ' Mise à jour de la chaîne de connexion avec les nouveaux paramètres
            DatabaseConnection.SetConnectionString(serveur, baseDeDonnees, authentificationWindows, utilisateur, motDePasse)
            
            ' Test de la connexion
            lblStatus.Text = "Test de la connexion en cours..."
            Application.DoEvents()
            
            If DatabaseConnection.TestConnection() Then
                lblStatus.Text = "Connexion réussie!"
                MessageBox.Show("Connexion à la base de données réussie.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                lblStatus.Text = "Échec de la connexion. Vérifiez les paramètres."
                MessageBox.Show("Impossible de se connecter à la base de données. Vérifiez les paramètres.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            lblStatus.Text = $"Erreur: {ex.Message}"
            MessageBox.Show($"Une erreur est survenue: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Enregistrer
    Private Sub BtnEnregistrer_Click(sender As Object, e As EventArgs)
        Try
            ' Récupération des valeurs saisies
            Dim serveur As String = txtServeur.Text.Trim()
            Dim baseDeDonnees As String = txtBaseDeDonnees.Text.Trim()
            Dim authentificationWindows As Boolean = chkAuthentificationWindows.Checked
            Dim utilisateur As String = txtUtilisateur.Text.Trim()
            Dim motDePasse As String = txtMotDePasse.Text
            
            ' Validation basique des entrées
            If String.IsNullOrEmpty(serveur) Then
                lblStatus.Text = "Veuillez saisir le nom du serveur."
                Return
            End If
            
            If String.IsNullOrEmpty(baseDeDonnees) Then
                lblStatus.Text = "Veuillez saisir le nom de la base de données."
                Return
            End If
            
            If Not authentificationWindows AndAlso String.IsNullOrEmpty(utilisateur) Then
                lblStatus.Text = "Veuillez saisir un nom d'utilisateur SQL."
                Return
            End If
            
            ' Mise à jour de la chaîne de connexion avec les nouveaux paramètres
            DatabaseConnection.SetConnectionString(serveur, baseDeDonnees, authentificationWindows, utilisateur, motDePasse)
            
            ' Sauvegarde des paramètres dans un fichier de configuration
            EnregistrerConfiguration(serveur, baseDeDonnees, authentificationWindows, utilisateur, motDePasse)
            
            DialogResult = DialogResult.OK
            Close()
        Catch ex As Exception
            lblStatus.Text = $"Erreur: {ex.Message}"
            MessageBox.Show($"Une erreur est survenue: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Annuler
    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
    
    ' Méthode pour enregistrer la configuration dans un fichier
    Private Sub EnregistrerConfiguration(serveur As String, baseDeDonnees As String, authentificationWindows As Boolean, utilisateur As String, motDePasse As String)
        Try
            Using writer As New StreamWriter(ConfigFile)
                writer.WriteLine($"Serveur={serveur}")
                writer.WriteLine($"BaseDeDonnees={baseDeDonnees}")
                writer.WriteLine($"AuthentificationWindows={authentificationWindows}")
                If Not authentificationWindows Then
                    writer.WriteLine($"Utilisateur={utilisateur}")
                    ' Noter que le mot de passe est stocké en clair. Pour une application de production,
                    ' il faudrait envisager un chiffrement plus sécurisé.
                    writer.WriteLine($"MotDePasse={motDePasse}")
                End If
            End Using
            
            lblStatus.Text = "Configuration enregistrée avec succès."
        Catch ex As Exception
            MessageBox.Show($"Erreur lors de l'enregistrement de la configuration: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Méthode pour charger la configuration depuis un fichier
    Private Sub ChargerConfiguration()
        Try
            If File.Exists(ConfigFile) Then
                Dim serveur As String = ""
                Dim baseDeDonnees As String = ""
                Dim authentificationWindows As Boolean = True
                Dim utilisateur As String = ""
                Dim motDePasse As String = ""
                
                Using reader As New StreamReader(ConfigFile)
                    Dim line As String
                    While Not reader.EndOfStream
                        line = reader.ReadLine()
                        If line.StartsWith("Serveur=") Then
                            serveur = line.Substring("Serveur=".Length)
                        ElseIf line.StartsWith("BaseDeDonnees=") Then
                            baseDeDonnees = line.Substring("BaseDeDonnees=".Length)
                        ElseIf line.StartsWith("AuthentificationWindows=") Then
                            Boolean.TryParse(line.Substring("AuthentificationWindows=".Length), authentificationWindows)
                        ElseIf line.StartsWith("Utilisateur=") Then
                            utilisateur = line.Substring("Utilisateur=".Length)
                        ElseIf line.StartsWith("MotDePasse=") Then
                            motDePasse = line.Substring("MotDePasse=".Length)
                        End If
                    End While
                End Using
                
                ' Mise à jour des contrôles avec les valeurs chargées
                txtServeur.Text = serveur
                txtBaseDeDonnees.Text = baseDeDonnees
                chkAuthentificationWindows.Checked = authentificationWindows
                txtUtilisateur.Text = utilisateur
                txtMotDePasse.Text = motDePasse
                
                ' Mise à jour de l'état des contrôles d'authentification SQL
                txtUtilisateur.Enabled = Not authentificationWindows
                txtMotDePasse.Enabled = Not authentificationWindows
                
                ' Mise à jour de la chaîne de connexion avec les valeurs chargées
                DatabaseConnection.SetConnectionString(serveur, baseDeDonnees, authentificationWindows, utilisateur, motDePasse)
                
                lblStatus.Text = "Configuration chargée."
            End If
        Catch ex As Exception
            MessageBox.Show($"Erreur lors du chargement de la configuration: {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Méthode statique pour afficher la boîte de dialogue de configuration
    Public Shared Function ShowConfigurationDialog() As Boolean
        Using form As New FormConnexion()
            Return form.ShowDialog() = DialogResult.OK
        End Using
    End Function
End Class