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
                
                ' Créer la commande SQL avec les paramètres basée sur la structure réelle de la table
                Dim query As String = "SELECT * FROM SPE_EXO WHERE OPE_NOOE = @NumeroCommande"
                
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
End Class