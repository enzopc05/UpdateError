Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class Cryptage
    ' Clé et vecteur d'initialisation pour le cryptage
    Private Shared ReadOnly _key As Byte() = Encoding.UTF8.GetBytes("Eurodislog2023!!")
    Private Shared ReadOnly _iv As Byte() = Encoding.UTF8.GetBytes("SecureEurodislog")

    ' Crypter une chaîne
    Public Shared Function CrypterTexte(texte As String) As String
        If String.IsNullOrEmpty(texte) Then
            Return texte
        End If
        
        Try
            Using aes As Aes = Aes.Create()
                aes.Key = _key
                aes.IV = _iv
                
                Dim crypteur As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)
                
                Using memoryStream As New MemoryStream()
                    Using cryptoStream As New CryptoStream(memoryStream, crypteur, CryptoStreamMode.Write)
                        Using writer As New StreamWriter(cryptoStream)
                            writer.Write(texte)
                        End Using
                    End Using
                    
                    Return Convert.ToBase64String(memoryStream.ToArray())
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine($"Erreur lors du cryptage: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' Décrypter une chaîne
    Public Shared Function DecrypterTexte(texteCrypte As String) As String
        If String.IsNullOrEmpty(texteCrypte) Then
            Return texteCrypte
        End If
        
        Try
            Using aes As Aes = Aes.Create()
                aes.Key = _key
                aes.IV = _iv
                
                Dim decrypteur As ICryptoTransform = aes.CreateDecryptor(aes.Key, aes.IV)
                
                Using memoryStream As New MemoryStream(Convert.FromBase64String(texteCrypte))
                    Using cryptoStream As New CryptoStream(memoryStream, decrypteur, CryptoStreamMode.Read)
                        Using reader As New StreamReader(cryptoStream)
                            Return reader.ReadToEnd()
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine($"Erreur lors du décryptage: {ex.Message}")
            Return String.Empty
        End Try
    End Function
End Class