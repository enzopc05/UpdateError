Imports Microsoft.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms

Public Class DatabaseConnection
    ' Chaîne de connexion par défaut (sera mise à jour via l'interface de configuration)
    Private Shared _connectionString As String = "Data Source=localhost;Initial Catalog=master;Integrated Security=True;TrustServerCertificate=True;"
    
    ' Méthode pour obtenir une connexion à la base de données
    Public Shared Function GetConnection() As SqlConnection
        Dim connection As New SqlConnection(_connectionString)
        Return connection
    End Function
    
    ' Méthode pour tester la connexion
    Public Shared Function TestConnection() As Boolean
        Try
            Using connection As SqlConnection = GetConnection()
                connection.Open()
                Return True
            End Using
        Catch ex As Exception
            Console.WriteLine($"Erreur de connexion à la base de données: {ex.Message}")
            Return False
        End Try
    End Function
    
    ' Méthode pour obtenir la chaîne de connexion
    Public Shared ReadOnly Property ConnectionString As String
        Get
            Return _connectionString
        End Get
    End Property
    
    ' Méthode pour définir une nouvelle chaîne de connexion (peut être utile pour une configuration dynamique)
    Public Shared Sub SetConnectionString(serverName As String, databaseName As String, useIntegratedSecurity As Boolean, Optional username As String = "", Optional password As String = "")
        Dim builder As New SqlConnectionStringBuilder With {
            .DataSource = serverName,
            .InitialCatalog = databaseName,
            .IntegratedSecurity = useIntegratedSecurity,
            .TrustServerCertificate = True
        }
        
        If Not useIntegratedSecurity Then
            builder.UserID = username
            builder.Password = password
        End If
        
        _connectionString = builder.ToString()
    End Sub
    
    ' Méthode pour charger la chaîne de connexion depuis le fichier de configuration
    Public Shared Sub LoadFromConfigFile()
        Dim configFile As String = Path.Combine(Application.StartupPath, "config.ini")
        
        If File.Exists(configFile) Then
            Dim serveur As String = ""
            Dim baseDeDonnees As String = ""
            Dim authentificationWindows As Boolean = True
            Dim utilisateur As String = ""
            Dim motDePasse As String = ""
            
            Using reader As New StreamReader(configFile)
                Dim line As String
                While Not reader.EndOfStream
                    line = reader.ReadLine()
                    If line.StartsWith("Serveur=") Then
                        serveur = Cryptage.DecrypterTexte(line.Substring("Serveur=".Length))
                    ElseIf line.StartsWith("BaseDeDonnees=") Then
                        baseDeDonnees = Cryptage.DecrypterTexte(line.Substring("BaseDeDonnees=".Length))
                    ElseIf line.StartsWith("AuthentificationWindows=") Then
                        Boolean.TryParse(line.Substring("AuthentificationWindows=".Length), authentificationWindows)
                    ElseIf line.StartsWith("Utilisateur=") Then
                        utilisateur = Cryptage.DecrypterTexte(line.Substring("Utilisateur=".Length))
                    ElseIf line.StartsWith("MotDePasse=") Then
                        motDePasse = Cryptage.DecrypterTexte(line.Substring("MotDePasse=".Length))
                    End If
                End While
            End Using
            
            ' Mettre à jour la chaîne de connexion avec les valeurs chargées
            SetConnectionString(serveur, baseDeDonnees, authentificationWindows, utilisateur, motDePasse)
        End If
    End Sub
End Class