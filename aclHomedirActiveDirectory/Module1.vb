Imports System.IO
Imports System.Text.RegularExpressions
Imports System.DirectoryServices
Imports System.Security.AccessControl

Module Module1
    ' Configuration AD 2000
    Private adresseActiveDirectory2000 As String = "LDAP://172.20.1.1/DC=Global,DC=ASP"
    Private login2000 As String = "admfma@global.asp"
    Private password2000 As String = ""

    ' Configuration AD 2003
    Private adresseActiveDirectory2003 As String = "LDAP://172.20.1.101/DC=GlobalSP,DC=local"
    Private login2003 As String = "ldap@globalsp.local"
    Private password2003 As String = ""

    ' Configuration commune 2000-2003
    'Private collectionHomedirAD As New Hashtable
    'Private collectionHomedirFichier As New Hashtable
    'Private sHomedirFichierProbleme As String = "\\STS01008\D$\Log\HomedirFichierProbleme.log"
    'Private sHomedirADProbleme As String = "\\STS01008\D$\Log\HomedirADProbleme.log"
    'Private sHomedirFichier As String = "\\STS01008\D$\Log\HomedirFichier.log"
    'Private sHomedirAD As String = "\\STS01008\D$\Log\HomedirAD.log"

    Sub Main()
        'Dim fsFIC As New System.IO.FileStream(sHomedirFichierProbleme, FileMode.Append, FileAccess.Write)
        'Dim fsAD As New System.IO.FileStream(sHomedirADProbleme, FileMode.Append, FileAccess.Write)
        'Dim swFIC As New System.IO.StreamWriter(fsFIC)
        'Dim swAD As New System.IO.StreamWriter(fsAD)
        'Dim messageMail As String = ""

        getHomedirProfileActiveDirectory(adresseActiveDirectory2000, login2000, password2000)
        getHomedirProfileActiveDirectory(adresseActiveDirectory2003, login2003, password2003)
        'messageMail = messageMail & "Dossiers existants sur les serveurs de fichier mais ne disposant pas d'entrées dans Active Directory :" & vbCrLf & vbCrLf

        'For Each membreCollectionHomedirFichier As String In collectionHomedirFichier.Values
        '    If Not collectionHomedirAD.ContainsValue(membreCollectionHomedirFichier) Then
        '        swFIC.WriteLine(membreCollectionHomedirFichier)
        '        messageMail = messageMail & membreCollectionHomedirFichier & vbCrLf
        '    End If
        '    swFIC.Flush()
        'Next
        'swFIC.Close()
        'fsFIC.Close()
        'messageMail = "Chemins étant rattachés à des utilisateurs Active Directory mais n'existant pas sur les serveurs de fichiers :" & vbCrLf & vbCrLf

        'For Each membreCollectionHomedirAD As String In collectionHomedirAD.Values
        '    If Not collectionHomedirFichier.ContainsValue(membreCollectionHomedirAD) And membreCollectionHomedirAD.IndexOf("\\FCH01001\ENTREPRISES$\ADMINGLOBAL\") = -1 Then
        '        swAD.WriteLine(membreCollectionHomedirAD)
        '        messageMail = messageMail & membreCollectionHomedirAD & vbCrLf
        '    End If
        '    swAD.Flush()
        'Next
        'swAD.Close()
        'swAD.Close()
    End Sub

    Private Sub getHomedirProfileActiveDirectory(ByVal adresseActiveDirectory, ByVal login, ByVal password)
        Dim monEntry As New DirectoryEntry(adresseActiveDirectory, login, password)
        Dim RechercheAll As New DirectorySearcher(monEntry)
        'Dim tmpHomedir As String

        RechercheAll.Filter = "(&(objectclass=user)(!objectclass=computer)(cn=*))"
        RechercheAll.PropertiesToLoad.Add("homeDirectory")
        RechercheAll.PropertiesToLoad.Add("profilePath")
        RechercheAll.PropertiesToLoad.Add("samAccountName")
        RechercheAll.PropertiesToLoad.Add("userAccountControl")
        RechercheAll.PageSize = 20000

        Dim myresult As SearchResultCollection = RechercheAll.FindAll()

        For i As Int32 = 0 To myresult.Count - 1
            If myresult.Item(i).Properties("homeDirectory").Count > 0 Then
                'tmpHomedir = myresult.Item(i).Properties("homeDirectory")(0).ToString.ToUpper
                'tmpHomedir = Replace(tmpHomedir, "_HOMEDIR", "\HOMEDIR")
                'tmpHomedir = Replace(tmpHomedir, "FIC01003\", "FIC01003\D$\")
                'collectionHomedir.Add(tmpHomedir, tmpHomedir)
            End If
            If myresult.Item(i).Properties("profilePath").Count > 0 Then
                'tmpHomedir = myresult.Item(i).Properties("profilePath")(0).ToString.ToUpper
                'tmpHomedir = Replace(tmpHomedir, "_PROFILES", "\PROFILES")
                'tmpHomedir = Replace(tmpHomedir, "FIC01003\", "FIC01003\D$\")
                'collectionHomedir.Add(tmpHomedir, tmpHomedir)
            End If
        Next
    End Sub

    Private Function IsAccountActive(ByVal userAccountControl As Integer) As Boolean





        Const UF_ACCOUNTDISABLE = 2
        'Dim accountDisabled As Integer = Convert.ToInt32(UF_ACCOUNTDISABLE)
        Dim flagExists As Integer = userAccountControl And UF_ACCOUNTDISABLE
        'if a match is found, then the disabled flag exists within the control flags
        If flagExists > 0 Then
            Return False
        Else
            Return True
        End If
    End Function
End Module


