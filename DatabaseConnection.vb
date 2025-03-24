Imports Microsoft.Data.SqlClient

Public Class DatabaseConnection
    ' Chaîne de connexion par défaut (sera mise à jour via l'interface de configuration)
    Private Shared _connectionString As String = "Data Source=localhost;Initial Catalog=master;Integrated Security=True;"
    
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
            .IntegratedSecurity = useIntegratedSecurity
        }
        
        If Not useIntegratedSecurity Then
            builder.UserID = username
            builder.Password = password
        End If
        
        _connectionString = builder.ToString()
    End Sub
End Class