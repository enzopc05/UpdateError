Imports System
Imports System.Windows.Forms
Imports System.IO

Module Program
    ' Point d'entrée principal pour l'application
    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        
        ' Vérifier si la connexion est déjà configurée
        Dim configFileExists As Boolean = File.Exists(Path.Combine(Application.StartupPath, "config.ini"))
        
        ' Si la configuration existe, essayer de la charger et tester la connexion
        If configFileExists Then
            Try
                ' Charger la configuration silencieusement (sans afficher le formulaire)
                DatabaseConnection.LoadFromConfigFile()
                
                ' Si la connexion échoue, afficher la boîte de dialogue de configuration
                If Not DatabaseConnection.TestConnection() Then
                    MessageBox.Show("La connexion à la base de données a échoué avec les paramètres sauvegardés." & vbCrLf & _
                                   "Veuillez vous connecter en tant qu'administrateur pour configurer la connexion.", _
                                   "Erreur de connexion", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    If Not FormLogin.ShowLoginDialog() Then
                        ' L'utilisateur a annulé l'authentification
                        Return
                    End If
                    
                    If Not FormConnexion.ShowConfigurationDialog() Then
                        ' L'utilisateur a annulé la configuration
                        MessageBox.Show("L'application ne peut pas fonctionner sans connexion à la base de données.", _
                                       "Configuration annulée", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                End If
            Catch ex As Exception
                ' En cas d'erreur de chargement, demander une authentification admin
                MessageBox.Show("Erreur lors du chargement de la configuration: " & ex.Message & vbCrLf & _
                               "Veuillez vous connecter en tant qu'administrateur pour configurer la connexion.", _
                               "Erreur de configuration", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                
                If Not FormLogin.ShowLoginDialog() Then
                    ' L'utilisateur a annulé l'authentification
                    Return
                End If
                
                If Not FormConnexion.ShowConfigurationDialog() Then
                    MessageBox.Show("L'application ne peut pas fonctionner sans connexion à la base de données.", _
                                   "Configuration annulée", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End Try
        Else
            ' Si aucune configuration n'existe, demander une authentification admin
            MessageBox.Show("Aucune configuration de connexion trouvée." & vbCrLf & _
                           "Veuillez vous connecter en tant qu'administrateur pour configurer la connexion.", _
                           "Configuration requise", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
            If Not FormLogin.ShowLoginDialog() Then
                ' L'utilisateur a annulé l'authentification
                Return
            End If
            
            If Not FormConnexion.ShowConfigurationDialog() Then
                MessageBox.Show("L'application ne peut pas fonctionner sans connexion à la base de données.", _
                               "Configuration annulée", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If
        
        ' Démarrer l'application principale
        Application.Run(New FormMain())
    End Sub
End Module