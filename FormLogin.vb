Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class FormLogin
    Inherits System.Windows.Forms.Form
    
    ' Déclaration des contrôles
    Private txtUtilisateur As TextBox
    Private txtMotDePasse As TextBox
    Private btnConnexion As Button
    Private btnAnnuler As Button
    Private lblStatus As Label
    
    ' Informations de connexion admin (normalement stockées de façon sécurisée)
    ' Note: Dans une application réelle, ces valeurs devraient être cryptées et/ou stockées de manière plus sécurisée
    Private ReadOnly _adminUser As String = "admin"
    Private ReadOnly _adminPassword As String = "adminsadmin"
    
    ' Couleur de la charte graphique Eurodislog
    Private ReadOnly _colorEurodislogBlue As Color = Color.FromArgb(0, 175, 215)  ' Bleu Eurodislog (#00AFD7)
    
    ' Constructeur
    Public Sub New()
        InitializeComponent()
    End Sub
    
    ' Initialisation des composants
    Private Sub InitializeComponent()
        ' Configuration de la fenêtre
        Me.Text = "Connexion Administrateur - Eurodislog"
        Me.Size = New System.Drawing.Size(400, 250)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        
        ' Création des labels
        Dim lblUtilisateur As New Label()
        lblUtilisateur.Text = "Utilisateur:"
        lblUtilisateur.Location = New System.Drawing.Point(50, 40)
        lblUtilisateur.Size = New System.Drawing.Size(100, 20)
        lblUtilisateur.Font = New Font("Segoe UI", 10)
        
        Dim lblMotDePasse As New Label()
        lblMotDePasse.Text = "Mot de passe:"
        lblMotDePasse.Location = New System.Drawing.Point(50, 80)
        lblMotDePasse.Size = New System.Drawing.Size(100, 20)
        lblMotDePasse.Font = New Font("Segoe UI", 10)
        
        ' Création des TextBox
        txtUtilisateur = New TextBox()
        txtUtilisateur.Location = New System.Drawing.Point(160, 40)
        txtUtilisateur.Size = New System.Drawing.Size(180, 20)
        txtUtilisateur.Font = New Font("Segoe UI", 10)
        
        txtMotDePasse = New TextBox()
        txtMotDePasse.Location = New System.Drawing.Point(160, 80)
        txtMotDePasse.Size = New System.Drawing.Size(180, 20)
        txtMotDePasse.PasswordChar = "*"c
        txtMotDePasse.Font = New Font("Segoe UI", 10)
        
        ' Création des boutons
        btnConnexion = New Button()
        btnConnexion.Text = "Connexion"
        btnConnexion.Location = New System.Drawing.Point(100, 130)
        btnConnexion.Size = New System.Drawing.Size(100, 30)
        btnConnexion.Font = New Font("Segoe UI", 10)
        btnConnexion.BackColor = _colorEurodislogBlue
        btnConnexion.ForeColor = Color.White
        btnConnexion.FlatStyle = FlatStyle.Flat
        btnConnexion.FlatAppearance.BorderSize = 0
        AddHandler btnConnexion.Click, AddressOf BtnConnexion_Click
        
        btnAnnuler = New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Location = New System.Drawing.Point(220, 130)
        btnAnnuler.Size = New System.Drawing.Size(100, 30)
        btnAnnuler.Font = New Font("Segoe UI", 10)
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        
        ' Label de statut
        lblStatus = New Label()
        lblStatus.Text = "Veuillez vous connecter en tant qu'administrateur."
        lblStatus.Location = New System.Drawing.Point(50, 180)
        lblStatus.Size = New System.Drawing.Size(300, 20)
        lblStatus.Font = New Font("Segoe UI", 9)
        
        ' Ajout des contrôles au formulaire
        Me.Controls.Add(lblUtilisateur)
        Me.Controls.Add(lblMotDePasse)
        Me.Controls.Add(txtUtilisateur)
        Me.Controls.Add(txtMotDePasse)
        Me.Controls.Add(btnConnexion)
        Me.Controls.Add(btnAnnuler)
        Me.Controls.Add(lblStatus)
        
        ' Définir AcceptButton et CancelButton pour simplifier la navigation au clavier
        Me.AcceptButton = btnConnexion
        Me.CancelButton = btnAnnuler
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Connexion
    Private Sub BtnConnexion_Click(sender As Object, e As EventArgs)
        Try
            ' Récupération des valeurs saisies
            Dim utilisateur As String = txtUtilisateur.Text.Trim()
            Dim motDePasse As String = txtMotDePasse.Text
            
            ' Validation des entrées
            If String.IsNullOrEmpty(utilisateur) OrElse String.IsNullOrEmpty(motDePasse) Then
                lblStatus.Text = "Veuillez remplir tous les champs."
                lblStatus.ForeColor = Color.Red
                Return
            End If
            
            ' Vérification des identifiants
            If utilisateur = _adminUser AndAlso motDePasse = _adminPassword Then
                DialogResult = DialogResult.OK
                Close()
            Else
                lblStatus.Text = "Identifiants incorrects. Veuillez réessayer."
                lblStatus.ForeColor = Color.Red
                txtMotDePasse.Text = ""
                txtMotDePasse.Focus()
            End If
        Catch ex As Exception
            lblStatus.Text = $"Erreur: {ex.Message}"
            lblStatus.ForeColor = Color.Red
        End Try
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Annuler
    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
    
    ' Méthode statique pour afficher la boîte de dialogue de connexion
    Public Shared Function ShowLoginDialog() As Boolean
        Using form As New FormLogin()
            Return form.ShowDialog() = DialogResult.OK
        End Using
    End Function
End Class