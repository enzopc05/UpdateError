Imports System.Data
Imports Microsoft.Data.SqlClient
Imports System.Collections.Generic

Public Class DataAccess
    ' Méthode pour récupérer les lignes de commande en fonction du numéro de commande et de la ligne de mission
    Public Shared Function GetLignesCommande(numeroCommande As String, numeroLigneMission As String) As DataTable
        Dim dataTable As New DataTable()
        
        Try
            Using connection As SqlConnection = DatabaseConnection.GetConnection()
                connection.Open()
                
                ' Requête SQL modifiée pour ne sélectionner que les colonnes souhaitées
                Dim query As String = "SELECT EXO_KEYU, OPE_NOOE, MIL_NOLM, EXO_SCAN, EXO_TRAK FROM SPE_EXO WHERE OPE_NOOE = @NumeroCommande"
                
                ' Ajouter la condition sur le numéro de ligne de mission si elle est fournie
                If Not String.IsNullOrEmpty(numeroLigneMission) Then
                    query &= " AND MIL_NOLM = @NumeroLigneMission"
                End If
                
                ' Ajouter le critère EXO_STAT<>'050'
                query &= " AND EXO_STAT<>'050'"
                
                Using command As New SqlCommand(query, connection)
                    ' Ajouter les paramètres pour éviter les injections SQL
                    command.Parameters.AddWithValue("@NumeroCommande", numeroCommande)
                    
                    If Not String.IsNullOrEmpty(numeroLigneMission) Then
                        command.Parameters.AddWithValue("@NumeroLigneMission", numeroLigneMission)
                    End If
                    
                    ' Exécuter la requête et remplir le DataTable
                    Using adapter As New SqlDataAdapter(command)
                        adapter.Fill(dataTable)
                    End Using
                End Using
                
                Return dataTable
            End Using
        Catch ex As Exception
            ' Gérer l'exception
            Console.WriteLine($"Erreur lors de la récupération des données: {ex.Message}")
            Throw
        End Try
    End Function
    
    ' Méthode pour vérifier si une ligne est en erreur selon les critères définis
    Public Shared Function EstEnErreur(row As DataRow) As Boolean
        Try
            ' Vérifier le critère sur la longueur de EXO_TRAK
            ' Si la longueur de EXO_TRAK n'est ni 9 ni 16, alors c'est une erreur
            If row.Table.Columns.Contains("EXO_TRAK") AndAlso Not row.IsNull("EXO_TRAK") Then
                Dim exoTrak As String = row("EXO_TRAK").ToString().Trim()
                If exoTrak.Length <> 9 AndAlso exoTrak.Length <> 16 Then
                    Return True
                End If
            End If
            
            ' Si aucune règle n'est violée, retourner false
            Return False
        Catch ex As Exception
            Console.WriteLine($"Erreur lors de la vérification des critères d'erreur: {ex.Message}")
            Return False
        End Try
    End Function
    
    ' Méthode pour mettre à jour la valeur d'EXO_TRAK
    Public Shared Function UpdateExoTrak(exoKeyU As String, nouvelleValeur As String) As Boolean
        Try
            Using connection As SqlConnection = DatabaseConnection.GetConnection()
                connection.Open()
                
                ' Créer la commande SQL avec les paramètres
                Dim query As String = "UPDATE SPE_EXO SET EXO_TRAK = @NouvelleValeur WHERE EXO_KEYU = @ExoKeyU"
                
                Using command As New SqlCommand(query, connection)
                    ' Ajouter les paramètres pour éviter les injections SQL
                    command.Parameters.AddWithValue("@NouvelleValeur", nouvelleValeur)
                    command.Parameters.AddWithValue("@ExoKeyU", exoKeyU)
                    
                    ' Exécuter la requête et retourner vrai si au moins une ligne a été mise à jour
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    Return rowsAffected > 0
                End Using
            End Using
        Catch ex As Exception
            ' Gérer l'exception
            Console.WriteLine($"Erreur lors de la mise à jour de EXO_TRAK: {ex.Message}")
            Throw
        End Try
    End Function
    
    ' Méthode pour vérifier si une valeur EXO_TRAK est valide
    Public Shared Function EstExoTrakValide(exoTrak As String) As Boolean
        If String.IsNullOrEmpty(exoTrak) Then Return False
        
        Dim valeurTrimmed As String = exoTrak.Trim()
        Return valeurTrimmed.Length = 9 OrElse valeurTrimmed.Length = 16
    End Function
End Class