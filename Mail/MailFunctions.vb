Imports System.Net.Mail
<Microsoft.VisualBasic.ComClass()> Public Class MailFunctions

    <System.Runtime.InteropServices.ComVisible(True)> _
    Public Function swiTest() As String
        Return "Dies ist ein Test"
    End Function

    <System.Runtime.InteropServices.ComVisible(True)> _
    Public Sub emailsenden(ByVal empfaenger As String, ByVal betreff As String, ByVal body As String, Optional ByVal pfadanhang As String = "", Optional ByVal kopie As String = "")

        Dim mail As New Schwindt.MAPI.MAPI

        mail.AddRecipientTo(empfaenger)
        If Not kopie = "" Then
            mail.AddRecipientCC(kopie)
        End If
        If Not pfadanhang = "" Then
            mail.AddAttachment(pfadanhang)
        End If

        'mail.SendMailDirect(betreff, body)
        mail.SendMailPopup(betreff, body)

    End Sub

End Class
