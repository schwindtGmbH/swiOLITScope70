Imports System.Web
Imports System.Net

Module Functions

    Public Function IgetVersionsinfo(ByVal Mode As Int32, ByVal Anwendung As String, ByVal Version As String, ByVal VersionBis As String) As String()

        Dim myWebClient As WebClient
        myWebClient = New WebClient

        Dim dlstring As String
        Dim myResult As String
        Dim sEintrag() As String
        Dim sFelder() As String

        Dim sVersion As String
        Dim sDatum As String
        Dim sBearbeiter As String
        Dim sAenderung As String
        Dim sResult() As String
        Dim bmyResult() As Byte
        Dim enc As System.Text.Encoding = System.Text.Encoding.Default


        dlstring = "http://to-officeline.de/Webservice.svc/Mode=" & Mode & "/Anwendung=" & Anwendung & "/Version=" & Version & "/VersionBis=" & VersionBis & ""

        myResult = myWebClient.DownloadString(dlstring)
        bmyResult = enc.GetBytes(myResult)
        myResult = (New System.Text.UTF8Encoding).GetString(bmyResult)

        ReDim sResult(0)
        If myResult = Chr(34) & "0" & Chr(34) Then GoTo keineEintraege

        myResult = Mid(myResult, 2, Len(myResult) - 3)

        sEintrag = Split(myResult, "|")

        ReDim sResult(sEintrag.GetLength(0) - 1)

        For i = 0 To sEintrag.GetLength(0) - 1
            sFelder = Split(sEintrag(i), ";")

            sVersion = sFelder(0)
            sDatum = sFelder(1)
            sBearbeiter = sFelder(2)
            sAenderung = sFelder(3)

            sResult(i) = sVersion & ";" & sDatum & ";" & sBearbeiter & ";" & sAenderung
        Next

keineEintraege:

        Return sResult

    End Function



End Module
