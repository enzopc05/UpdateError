Imports System
Imports System.Windows.Forms

' IMPORTANT: Suppression du namespace pour que la méthode Main soit au niveau module
Module Program
    ' Point d'entrée principal pour l'application
    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        
        ' Afficher la boîte de dialogue de configuration si aucune connexion n'est disponible
        If Not DatabaseConnection.TestConnection() Then
            If Not FormConnexion.ShowConfigurationDialog() Then
                ' L'utilisateur a annulé la configuration
                MessageBox.Show("L'application ne peut pas fonctionner sans connexion à la base de données.", "Configuration annulée", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' Tester à nouveau la connexion
            If Not DatabaseConnection.TestConnection() Then
                MessageBox.Show("Impossible de se connecter à la base de données. Veuillez vérifier les paramètres de connexion.", "Erreur de connexion", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
        End If
        
        ' Démarrer l'application
        Application.Run(New FormMain())
    End Sub
End Module