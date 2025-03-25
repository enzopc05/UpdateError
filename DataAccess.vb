Imports System.Data
Imports Microsoft.Data.SqlClient
Imports System.Collections.Generic

Public Class DataAccess
    ' Méthode pour récupérer les lignes de commande en fonction du numéro de commande et de la ligne de mission
    Public Shared Function GetLignesCommande(numeroCommande As String, numeroLigneMission As String) As DataTable
        Dim resultats As New DataTable()
        
        Try
            Using connection As SqlConnection = DatabaseConnection.GetConnection()
                connection.Open()
                
                ' Requête SQL pour afficher uniquement les colonnes souhaitées
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
                        adapter.Fill(resultats)
                    End Using
                End Using
                
                ' Marquer les lignes avec l'information d'erreur
                MarquerLignesEnErreur(resultats, connection, numeroCommande, numeroLigneMission)
                
                Return resultats
            End Using
        Catch ex As Exception
            ' Gérer l'exception
            Console.WriteLine($"Erreur lors de la récupération des données: {ex.Message}")
            Throw
        End Try
    End Function
    
    ' Méthode pour marquer les lignes en erreur sans ajouter la colonne EXO_QTE au résultat
    Private Shared Sub MarquerLignesEnErreur(resultats As DataTable, connection As SqlConnection, numeroCommande As String, numeroLigneMission As String)
        Try
            ' Ajouter une colonne pour stocker l'information d'erreur et le message associé
            If Not resultats.Columns.Contains("EstEnErreur") Then
                resultats.Columns.Add("EstEnErreur", GetType(Boolean))
            End If
            
            If Not resultats.Columns.Contains("MessageErreur") Then
                resultats.Columns.Add("MessageErreur", GetType(String))
            End If
            
            ' Créer une requête pour récupérer les informations de quantité
            Dim query As String = "SELECT EXO_KEYU, EXO_QTE FROM SPE_EXO WHERE OPE_NOOE = @NumeroCommande"
            
            ' Ajouter la condition sur le numéro de ligne de mission si elle est fournie
            If Not String.IsNullOrEmpty(numeroLigneMission) Then
                query &= " AND MIL_NOLM = @NumeroLigneMission"
            End If
            
            ' Ajouter le critère EXO_STAT<>'050'
            query &= " AND EXO_STAT<>'050'"
            
            ' Créer un dictionnaire pour stocker les quantités
            Dim quantites As New Dictionary(Of String, Decimal)()
            
            Using command As New SqlCommand(query, connection)
                ' Ajouter les paramètres pour éviter les injections SQL
                command.Parameters.AddWithValue("@NumeroCommande", numeroCommande)
                
                If Not String.IsNullOrEmpty(numeroLigneMission) Then
                    command.Parameters.AddWithValue("@NumeroLigneMission", numeroLigneMission)
                End If
                
                ' Exécuter la requête et récupérer les quantités
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim exoKeyU As String = reader("EXO_KEYU").ToString()
                        Dim speQte As Decimal = 0
                        
                        If Not reader.IsDBNull(reader.GetOrdinal("EXO_QTE")) Then
                            Decimal.TryParse(reader("EXO_QTE").ToString(), speQte)
                        End If
                        
                        quantites(exoKeyU) = speQte
                    End While
                End Using
            End Using
            
            ' Parcourir les résultats et marquer les lignes en erreur
            For Each row As DataRow In resultats.Rows
                Dim exoKeyU As String = row("EXO_KEYU").ToString()
                Dim exoTrak As String = row("EXO_TRAK").ToString().Trim()
                Dim speQte As Decimal = 0
                
                ' Récupérer la quantité pour cette ligne
                If quantites.ContainsKey(exoKeyU) Then
                    speQte = quantites(exoKeyU)
                End If
                
                ' Vérifier si la ligne est en erreur selon les critères définis
                Dim enErreur As Boolean = False
                Dim messageErreur As String = ""
                
                ' Si EXO_QTE > 1, EXO_TRAK doit faire 16 caractères
                If speQte > 1 AndAlso exoTrak.Length <> 16 Then
                    enErreur = True
                    messageErreur = "La quantité est supérieure à 1, le numéro de tracking doit comporter exactement 16 caractères."
                ' Si EXO_QTE = 1, EXO_TRAK doit faire 9 caractères
                ElseIf speQte = 1 AndAlso exoTrak.Length <> 9 Then
                    enErreur = True
                    messageErreur = "La quantité est 1, le numéro de tracking doit comporter exactement 9 caractères."
                End If
                
                ' Marquer la ligne
                row("EstEnErreur") = enErreur
                row("MessageErreur") = messageErreur
            Next
        Catch ex As Exception
            Console.WriteLine($"Erreur lors du marquage des lignes en erreur: {ex.Message}")
        End Try
    End Sub
    
    ' Méthode pour vérifier si une ligne est en erreur selon les critères définis
    Public Shared Function EstEnErreur(row As DataRow) As Boolean
        Try
            If row.Table.Columns.Contains("EstEnErreur") Then
                Return CBool(row("EstEnErreur"))
            End If
            
            Return False
        Catch ex As Exception
            Console.WriteLine($"Erreur lors de la vérification des critères d'erreur: {ex.Message}")
            Return False
        End Try
    End Function
    
    ' Méthode pour obtenir le message d'erreur pour une ligne
    Public Shared Function GetMessageErreur(row As DataRow) As String
        Try
            If row.Table.Columns.Contains("MessageErreur") AndAlso Not row.IsNull("MessageErreur") Then
                Return row("MessageErreur").ToString()
            End If
            
            Return ""
        Catch ex As Exception
            Console.WriteLine($"Erreur lors de la récupération du message d'erreur: {ex.Message}")
            Return ""
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
    
    ' Méthode pour vérifier si une valeur EXO_TRAK est valide selon la quantité
    Public Shared Function EstExoTrakValide(exoTrak As String, speQte As Decimal) As Boolean
        If String.IsNullOrEmpty(exoTrak) Then Return False
        
        Dim valeurTrimmed As String = exoTrak.Trim()
        
        ' Si EXO_QTE > 1, EXO_TRAK doit faire 16 caractères
        If speQte > 1 Then
            Return valeurTrimmed.Length = 16
        ' Si EXO_QTE = 1, EXO_TRAK doit faire 9 caractères
        ElseIf speQte = 1 Then
            Return valeurTrimmed.Length = 9
        Else
            ' Si on ne peut pas déterminer la quantité, on applique la règle générale
            Return valeurTrimmed.Length = 9 OrElse valeurTrimmed.Length = 16
        End If
    End Function
    
    ' Récupérer la quantité pour un EXO_KEYU donné
    Public Shared Function GetQuantiteForExoKeyU(exoKeyU As String) As Decimal
        Try
            Using connection As SqlConnection = DatabaseConnection.GetConnection()
                connection.Open()
                
                ' Créer la commande SQL avec les paramètres
                Dim query As String = "SELECT EXO_QTE FROM SPE_EXO WHERE EXO_KEYU = @ExoKeyU"
                
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@ExoKeyU", exoKeyU)
                    
                    ' Exécuter la requête et récupérer la valeur
                    Dim result As Object = command.ExecuteScalar()
                    
                    If result IsNot Nothing AndAlso Not Convert.IsDBNull(result) Then
                        Dim speQte As Decimal = 0
                        If Decimal.TryParse(result.ToString(), speQte) Then
                            Return speQte
                        End If
                    End If
                    
                    ' Valeur par défaut si aucun résultat n'est trouvé
                    Return 0
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine($"Erreur lors de la récupération de la quantité: {ex.Message}")
            Return 0
        End Try
    End Function
End Class