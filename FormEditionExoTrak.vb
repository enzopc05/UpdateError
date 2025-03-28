Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class FormEditionExoTrak
    Inherits Form
    
    ' Propriétés pour stocker les données de la ligne sélectionnée
    Public Property ExoKeyU As String
    Public Property ExoTrak As String
    Public Property SpeQte As Decimal = 0
    
    ' Contrôles du formulaire
    Private lblExplication As Label
    Private lblExoKeyU As Label
    Private txtExoKeyU As TextBox
    Private lblSpeQte As Label
    Private txtSpeQte As TextBox
    Private lblExoTrak As Label
    Private txtExoTrak As TextBox
    Private lblRegles As Label
    Private lblCompteur As Label
    Private btnEnregistrer As Button
    Private btnAnnuler As Button
    
    ' Couleurs de la charte graphique
    Private ReadOnly colorBlue As Color = Color.FromArgb(0, 175, 215)  ' Bleu (#00AFD7)
    Private ReadOnly colorDarkBlue As Color = Color.FromArgb(47, 54, 127)  ' Bleu foncé (#2F367F)
    
    ' Constructeur
    Public Sub New()
        InitializeComponent()
    End Sub
    
    ' Initialisation des composants
    Private Sub InitializeComponent()
        ' Configuration de la fenêtre
        Me.Text = "Modification du numéro de tracking"
        Me.Size = New Size(450, 350)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.White
        Me.Font = New Font("Segoe UI", 9)
        
        ' Label d'explication
        lblExplication = New Label()
        lblExplication.Text = "Veuillez corriger le numéro de tracking pour qu'il respecte les règles de validation."
        lblExplication.Location = New Point(20, 20)
        lblExplication.Size = New Size(400, 40)
        lblExplication.Font = New Font("Segoe UI", 10)
        lblExplication.TextAlign = ContentAlignment.MiddleLeft
        
        ' Label pour ExoKeyU
        lblExoKeyU = New Label()
        lblExoKeyU.Text = "Identifiant de ligne:"
        lblExoKeyU.Location = New Point(20, 70)
        lblExoKeyU.Size = New Size(150, 20)
        lblExoKeyU.Font = New Font("Segoe UI", 9)
        
        ' TextBox pour ExoKeyU (en lecture seule)
        txtExoKeyU = New TextBox()
        txtExoKeyU.Location = New Point(170, 70)
        txtExoKeyU.Size = New Size(250, 25)
        txtExoKeyU.ReadOnly = True
        txtExoKeyU.BackColor = Color.FromArgb(240, 240, 240)
        
        ' Label pour SpeQte
        lblSpeQte = New Label()
        lblSpeQte.Text = "Quantité:"
        lblSpeQte.Location = New Point(20, 105)
        lblSpeQte.Size = New Size(150, 20)
        lblSpeQte.Font = New Font("Segoe UI", 9)
        
        ' TextBox pour SpeQte (en lecture seule)
        txtSpeQte = New TextBox()
        txtSpeQte.Location = New Point(170, 105)
        txtSpeQte.Size = New Size(250, 25)
        txtSpeQte.ReadOnly = True
        txtSpeQte.BackColor = Color.FromArgb(240, 240, 240)
        
        ' Label pour ExoTrak
        lblExoTrak = New Label()
        lblExoTrak.Text = "Numéro de tracking:"
        lblExoTrak.Location = New Point(20, 140)
        lblExoTrak.Size = New Size(150, 20)
        lblExoTrak.Font = New Font("Segoe UI", 9)
        
        ' TextBox pour ExoTrak
        txtExoTrak = New TextBox()
        txtExoTrak.Location = New Point(170, 140)
        txtExoTrak.Size = New Size(250, 25)

        ' Label pour le compteur de caractères
        lblCompteur = New Label()
        lblCompteur.Text = "0 caractère(s)"
        lblCompteur.Location = New Point(170, 165)
        lblCompteur.Size = New Size(250, 20)
        lblCompteur.Font = New Font("Segoe UI", 8)
        lblCompteur.ForeColor = Color.Gray
        
        ' Label pour les règles
        lblRegles = New Label()
        lblRegles.Text = "Le numéro de tracking doit respecter les règles suivantes :" & Environment.NewLine & _
                         "- 9 caractères si la quantité est égale à 1" & Environment.NewLine & _
                         "- 16 caractères si la quantité est supérieure à 1"
        lblRegles.Location = New Point(20, 195)
        lblRegles.Size = New Size(400, 60)
        lblRegles.Font = New Font("Segoe UI", 9, FontStyle.Italic)
        lblRegles.ForeColor = Color.Gray
        
        ' Bouton Enregistrer
        btnEnregistrer = New Button()
        btnEnregistrer.Text = "Enregistrer"
        btnEnregistrer.Location = New Point(120, 250)
        btnEnregistrer.Size = New Size(120, 35)
        btnEnregistrer.BackColor = colorBlue
        btnEnregistrer.ForeColor = Color.White
        btnEnregistrer.FlatStyle = FlatStyle.Flat
        btnEnregistrer.FlatAppearance.BorderSize = 0
        btnEnregistrer.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        AddHandler btnEnregistrer.Click, AddressOf BtnEnregistrer_Click
        
        ' Effet de survol pour le bouton Enregistrer
        AddHandler btnEnregistrer.MouseEnter, Sub(sender As Object, e As EventArgs)
            btnEnregistrer.BackColor = colorDarkBlue
        End Sub
        
        AddHandler btnEnregistrer.MouseLeave, Sub(sender As Object, e As EventArgs)
            btnEnregistrer.BackColor = colorBlue
        End Sub
        
        ' Bouton Annuler
        btnAnnuler = New Button()
        btnAnnuler.Text = "Annuler"
        btnAnnuler.Location = New Point(250, 250)
        btnAnnuler.Size = New Size(120, 35)
        btnAnnuler.BackColor = Color.FromArgb(240, 240, 240)
        btnAnnuler.ForeColor = Color.Black
        btnAnnuler.FlatStyle = FlatStyle.Flat
        btnAnnuler.FlatAppearance.BorderSize = 1
        btnAnnuler.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200)
        btnAnnuler.Font = New Font("Segoe UI", 10)
        AddHandler btnAnnuler.Click, AddressOf BtnAnnuler_Click
        
        ' Effet de survol pour le bouton Annuler
        AddHandler btnAnnuler.MouseEnter, Sub(sender As Object, e As EventArgs)
            btnAnnuler.BackColor = Color.FromArgb(230, 230, 230)
        End Sub
        
        AddHandler btnAnnuler.MouseLeave, Sub(sender As Object, e As EventArgs)
            btnAnnuler.BackColor = Color.FromArgb(240, 240, 240)
        End Sub
        
        ' Ajout des contrôles au formulaire
        Me.Controls.Add(lblExplication)
        Me.Controls.Add(lblExoKeyU)
        Me.Controls.Add(txtExoKeyU)
        Me.Controls.Add(lblSpeQte)
        Me.Controls.Add(txtSpeQte)
        Me.Controls.Add(lblExoTrak)
        Me.Controls.Add(txtExoTrak)
        Me.Controls.Add(lblRegles)
        Me.Controls.Add(lblCompteur)
        Me.Controls.Add(btnEnregistrer)
        Me.Controls.Add(btnAnnuler)
        
        ' Validation des entrées lorsque le texte change
        AddHandler txtExoTrak.TextChanged, AddressOf TxtExoTrak_TextChanged
    End Sub
    
' Méthode appelée lors du chargement du formulaire
Protected Overrides Sub OnLoad(e As EventArgs)
    MyBase.OnLoad(e)
    
    ' Remplir les champs avec les valeurs existantes
    txtExoKeyU.Text = ExoKeyU
    txtSpeQte.Text = SpeQte.ToString()
    txtExoTrak.Text = ExoTrak
    
    ' Vérifier si la valeur actuelle est valide
    ValidateExoTrak()
    
    ' Ajouter un gestionnaire pour l'événement Shown
    AddHandler Me.Shown, AddressOf FormEditionExoTrak_Shown
End Sub

' Gestionnaire d'événement qui s'exécute lorsque le formulaire est entièrement affiché
Private Sub FormEditionExoTrak_Shown(sender As Object, e As EventArgs)
    ' Donner le focus au champ de numéro de tracking
    txtExoTrak.Focus()
    
    ' Sélectionner tout le texte
    txtExoTrak.SelectionStart = 0
    txtExoTrak.SelectionLength = txtExoTrak.Text.Length
    
    ' Alternative si la méthode ci-dessus ne fonctionne pas
    txtExoTrak.Select(0, txtExoTrak.Text.Length)
End Sub
    ' Gestionnaire d'événement pour la modification du texte dans txtExoTrak
Private Sub TxtExoTrak_TextChanged(sender As Object, e As EventArgs)
    ' Mettre à jour le compteur de caractères
    Dim nombreCaracteres As Integer = txtExoTrak.Text.Trim().Length
    Dim texteCompteur As String = $"{nombreCaracteres} caractère(s)"
    
    ' Adapter la couleur en fonction de la validité
    If (SpeQte = 1 AndAlso nombreCaracteres = 9) OrElse (SpeQte > 1 AndAlso nombreCaracteres = 16) Then
        lblCompteur.ForeColor = Color.FromArgb(0, 150, 0)  ' Vert
        lblCompteur.Font = New Font(lblCompteur.Font, FontStyle.Bold)
    ElseIf nombreCaracteres > 0 Then
        lblCompteur.ForeColor = Color.FromArgb(213, 0, 0)  ' Rouge
        lblCompteur.Font = New Font(lblCompteur.Font, FontStyle.Bold)
    Else
        lblCompteur.ForeColor = Color.Gray
        lblCompteur.Font = New Font(lblCompteur.Font, FontStyle.Regular)
    End If
    
    lblCompteur.Text = texteCompteur
    
    ' Valider le champ
    ValidateExoTrak()
End Sub
    
    ' Méthode pour valider la valeur d'ExoTrak
    Private Sub ValidateExoTrak()
        Dim valeurActuelle As String = txtExoTrak.Text.Trim()
        Dim estValide As Boolean = DataAccess.EstExoTrakValide(valeurActuelle, SpeQte)
        
        ' Mettre à jour l'apparence du champ et du bouton Enregistrer
        If estValide Then
            txtExoTrak.BackColor = Color.FromArgb(240, 255, 240)  ' Vert très clair
            lblRegles.ForeColor = Color.FromArgb(0, 150, 0)  ' Vert
            btnEnregistrer.Enabled = True
        Else
            txtExoTrak.BackColor = Color.FromArgb(255, 235, 235)  ' Rouge très clair
            lblRegles.ForeColor = Color.FromArgb(213, 0, 0)  ' Rouge
            btnEnregistrer.Enabled = False
        End If
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Enregistrer
    Private Sub BtnEnregistrer_Click(sender As Object, e As EventArgs)
        Try
            Dim nouveauTrak As String = txtExoTrak.Text.Trim()
            
            ' Vérifier à nouveau si la valeur est valide
            If Not DataAccess.EstExoTrakValide(nouveauTrak, SpeQte) Then
                Dim message As String
                If SpeQte = 1 Then
                    message = "Le numéro de tracking doit comporter exactement 9 caractères car la quantité est égale à 1."
                ElseIf SpeQte > 1 Then
                    message = "Le numéro de tracking doit comporter exactement 16 caractères car la quantité est supérieure à 1."
                Else
                    message = "Le numéro de tracking doit comporter exactement 9 ou 16 caractères selon la quantité."
                End If
                
                MessageBox.Show(message, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' Mettre à jour la base de données
            If DataAccess.UpdateExoTrak(ExoKeyU, nouveauTrak) Then
                DialogResult = DialogResult.OK
                Close()
            Else
                MessageBox.Show("Aucune modification n'a été effectuée. La ligne n'a pas pu être mise à jour.", 
                               "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erreur lors de la mise à jour: {ex.Message}", 
                           "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ' Gestionnaire d'événement pour le bouton Annuler
    Private Sub BtnAnnuler_Click(sender As Object, e As EventArgs)
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
    
    ' Méthode statique pour afficher le formulaire d'édition
    Public Shared Function EditExoTrak(exoKeyU As String, exoTrak As String) As Tuple(Of Boolean, String)
        Using form As New FormEditionExoTrak()
            form.ExoKeyU = exoKeyU
            form.ExoTrak = exoTrak
            
            ' Récupérer la quantité pour cette ligne
            form.SpeQte = DataAccess.GetQuantiteForExoKeyU(exoKeyU)
            
            If form.ShowDialog() = DialogResult.OK Then
                ' Retourner vrai et la nouvelle valeur
                Return New Tuple(Of Boolean, String)(True, form.txtExoTrak.Text.Trim())
            Else
                ' Retourner faux et la valeur originale
                Return New Tuple(Of Boolean, String)(False, exoTrak)
            End If
        End Using
    End Function
End Class