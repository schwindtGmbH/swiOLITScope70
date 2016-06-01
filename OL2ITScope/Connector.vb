Imports System.Runtime.InteropServices
Imports Sagede.OfficeLine.Engine

<Microsoft.VisualBasic.ComClass()> Public Class Connector

    Public Shared _mandant As Mandant

    <System.Runtime.InteropServices.ComVisible(True)> _
    Public Function swiTest() As String
        swiTest = "Test MG"
    End Function

    Public Sub New()
        'MyBase.New()
    End Sub

    Public Sub InitMandant(ByVal Mandant As OLSysIInterop70.Mandant)
        If _mandant Is Nothing Then _mandant = DirectCast(Mandant, Sagede.OfficeLine.Interop70.Mandant).GetRealObject
    End Sub

    Public Sub InitMandantNET(Mandant As Mandant)
        If _mandant Is Nothing Then _mandant = Mandant
    End Sub

    Public Function CheckLizenz(ByVal Version As String) As Boolean
        If ICheckLizenz(Version) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckLizenzNewUpdate(ByVal Version As String) As Boolean
        If ICheckLizenzNewUpdate(Version) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckLizenzOffline() As Boolean
        If ICheckLizenzOffline() Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DecodeString(Text As String) As String
        Return IDecodeString(Text, "schwindt")
    End Function

    Public Function EncodeString(Text As String) As String
        Return IEncodeString(Text, "schwindt")
    End Function

    Public Sub ReadProductDatasheetByEAN(ean As String)
        IReadProductDatasheetByEAN(ean)
    End Sub

    Public Function CreateArticleEAN(ean As String) As String
        Return ICreateArticleEAN(ean)
    End Function

    Public Function CreateArticlePuid(puid As String, Optional Silent As Boolean = False) As CProduct
        Return ICreateArticlePuid(puid, Silent)
    End Function

    Public Function StrCreateArticlePuid(puid As String, CreateMessage As Boolean, Optional Silent As Boolean = False, Optional Shopware As Boolean = False) As String
        Return IStrCreateArticlePuid(puid, CreateMessage, Silent, Shopware)
    End Function

    Public Function CreateLoginINI(ByVal Benutzer As String, ByVal Kennwort As String)
        ICreateLoginINI(Benutzer, Kennwort)
    End Function

    Public Function ReadProductSets() As Boolean
        Return IReadProductSets()
    End Function

    Public Function ReadCompanyList() As Boolean
        Return IReadCompanyList()
    End Function

    Public Function MailProductDatasheetByEAN(ean As String)
        IMailProductDatasheetByEAN(ean)
    End Function

    Public Function UpdateArticlePuid(artikel As String, puid As String) As Boolean
        Return IUpdateArticlePuid(artikel, puid)
    End Function

    Public Sub CheckPuid4Article(article As String)
        ICheckPuid4Article(article)
    End Sub

    Public Function UpdateArticleFull(article As String) As Boolean
        Return IUpdateArticleFull(article)
    End Function

End Class
