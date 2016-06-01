Imports System.Net
Imports Sagede.Core.Tools

Module License

    Public Function ICheckLizenz(ByVal Version As String) As Boolean

        Dim KundenNr As String
        Dim Funktion As String
        Dim genericConn As Sagede.OfficeLine.Data.IGenericConnection

        genericConn = Connector._mandant.MainDevice.GenericConnection
        Dim lookup As New Sagede.OfficeLine.Data.GenericDataLookup(genericConn)

        KundenNr = Connector._mandant.License.SerialNumber
        If Len(KundenNr) < 9 Then KundenNr = "0" & KundenNr

        If CheckLizenzupdate(KundenNr, False, Version) = False Then
            Return False
        End If

        Funktion = ConversionHelper.ToString(lookup.GetString("Bezeichnung", "KHKGruppen", "Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"))

        If Funktion = "" Then
            MsgBox("Sie haben keine lizensierte Version!" & vbCrLf & "Bitte kontaktieren Sie Ihren Fachhändler!", vbCritical, "Keine Lizenz")
            Return False
        End If

        Funktion = IDecodeString(Funktion, "schwindt")

        If Funktion.Contains("Demo") = False Then
            Funktion = vb6_right(Funktion, 1)

            If Funktion = "1" Then
                Return True
            Else
                MsgBox("Sie haben keine lizensierte Version!" & vbCrLf & "Bitte kontaktieren Sie Ihren Fachhändler!", vbCritical, "Keine Lizenz")
                Return False
            End If
        Else
            Dim Datum As Date

            Datum = Mid(Funktion, 12, Len(Funktion) - 12)

            If (Datum - GetLizenzdatum()).Days < 0 Then
                MsgBox("Die Demoversion ist abgelaufen, bitte kontaktieren Sie Ihren Fachhändler!", vbCritical, "Lizenz abgelaufen")
                Return False
            Else
                MsgBox("Dies ist eine Demoversion, lauffähig bis zum " & Datum & "!", vbInformation, "Demoversion")
                Return True
            End If
        End If

        Return True

    End Function

    Public Function ICheckLizenzNewUpdate(ByVal Version As String) As Boolean

        Dim KundenNr As String
        Dim Funktion As String
        Dim genericConn As Sagede.OfficeLine.Data.IGenericConnection

        genericConn = Connector._mandant.MainDevice.GenericConnection
        Dim lookup As New Sagede.OfficeLine.Data.GenericDataLookup(genericConn)

        KundenNr = Connector._mandant.License.SerialNumber
        If Len(KundenNr) < 9 Then KundenNr = "0" & KundenNr

        If CheckLizenzupdate(KundenNr, True, Version) = False Then
            Return False
        End If

        Funktion = ConversionHelper.ToString(lookup.GetString("Bezeichnung", "KHKGruppen", "Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"))

        If Funktion = "" Then
            MsgBox("Sie haben keine lizensierte Version!" & vbCrLf & "Bitte kontaktieren Sie Ihren Fachhändler!", vbCritical, "Keine Lizenz")
            Return False
        End If

        Funktion = IDecodeString(Funktion, "schwindt")
        Funktion = vb6_right(Funktion, 1)

        If Funktion = "1" Then
            Return True
        Else
            MsgBox("Sie haben keine lizensierte Version!" & vbCrLf & "Bitte kontaktieren Sie Ihren Fachhändler!", vbCritical, "Keine Lizenz")
            Return False
        End If

        Return True

    End Function

    Public Function ICheckLizenzOffline() As Boolean

        Dim Funktion As String
        Dim genericConn As Sagede.OfficeLine.Data.IGenericConnection

        genericConn = Connector._mandant.MainDevice.GenericConnection
        Dim lookup As New Sagede.OfficeLine.Data.GenericDataLookup(genericConn)

        Funktion = ConversionHelper.ToString(lookup.GetString("Bezeichnung", "KHKGruppen", "Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"))

        If Funktion = "" Then
            MsgBox("Es wurden keine ITscope-Lizenzinformationen in der Datenbank gefunden!", MsgBoxStyle.Exclamation, "Fehler")
            Return False
        End If

        Funktion = IDecodeString(Funktion, "schwindt")
        Funktion = vb6_right(Funktion, 1)

        If Funktion = "1" Then
            Return True
        Else
            Return False
        End If

        Return True

    End Function

    Public Function CheckLizenzupdate(ByVal KundenNr As String, NewUpdate As Boolean, ByVal Version As String) As Boolean

        Dim Datum As Date = GetServerDate()

        If NewUpdate Then
            LizenzUpdate(KundenNr, Version)
        Else
            If GetLizenzdatum() = "01.01.2014" Then
                MsgBox("Sie müssen die Lizenz zuerst innerhalb der Versioninfo aktivieren!", MsgBoxStyle.Exclamation, "Lizenz-Aktivierung")
                Return False
            End If
            If (Datum - GetLizenzdatum()).Days >= 1 Then
                LizenzUpdate(KundenNr, Version)
            End If
        End If

        Return True

    End Function

    Public Sub LizenzUpdate(ByVal KundenNr As String, ByVal Version As String)

        Dim myWebClient As WebClient
        myWebClient = New WebClient

        Dim myResultEncode As String
        Dim myResult As String
        Dim dlString As String
        Dim sQry As String

        Dim MaxVersion As String

        dlString = "http://to-officeline.de/Webservice.svc/KundenNr=" & KundenNr & "/Anwendung=ITS"

        MaxVersion = GetMaxVersion(KundenNr)

        Try
            If MaxVersion <> "0" Then
                If CInt(Mid(Version, 1, 1)) <> CInt(Mid(MaxVersion, 1, 1)) Then GoTo AblaufDemo
                If CInt(Mid(Version, 3, 1)) <> CInt(Mid(MaxVersion, 3, 1)) Then GoTo AblaufDemo
                If CInt(Mid(Version, 5)) > CInt(Mid(MaxVersion, 5)) Then GoTo AblaufDemo
            End If

            myResult = myWebClient.DownloadString(dlString)
            myResult = Mid(myResult, 2, Len(myResult) - 2)

            myResultEncode = IEncodeString("ITScope" & myResult, "schwindt")

            sQry = "UPDATE KHKGruppen SET Bezeichnung='" & ConversionHelper.ToString(myResultEncode) & "' WHERE Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            sQry = "UPDATE KHKGruppen SET Bezeichnung='" & IEncodeString(GetServerDate(), "schwindt") & "' WHERE Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='License'"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            If myResult.Contains("Demo") = False Then
                If myResult = "1" Then
                    MsgBox("Lizenz erfolgreich abgerufen!", MsgBoxStyle.Information, "ITScope-Lizenz")
                Else
                    MsgBox("Das Produkt ist nicht lizenziert!", MsgBoxStyle.Exclamation, "ITScope-Lizenz")
                End If
            Else
                MsgBox("Lizenz erfolgreich abgerufen!" & vbNewLine & "Dies ist eine Demoversion, lauffähig bis zum " & Mid(myResult, 6, Len(myResult) - 6) & "!", vbInformation, "ITScope-Demoversion")
            End If
        Catch ex As Exception
            sQry = "UPDATE KHKGruppen SET Bezeichnung='" & IEncodeString(ConversionHelper.ToString(KundenNr & "0000ITS"), "schwindt") & "' WHERE Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

            MsgBox("Es ist ein Fehler beim Lizenzabruf aufgetreten!", MsgBoxStyle.Critical, "Lizenz")
        End Try

        Exit Sub

AblaufDemo:
        sQry = "UPDATE KHKGruppen SET Bezeichnung='" & IEncodeString(ConversionHelper.ToString(KundenNr & "0000ITS"), "schwindt") & "' WHERE Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='Funktionen'"
        Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(sQry)

        MsgBox("Diese Version ist für Sie nicht verfügbar!" & vbNewLine & "Bitte die bisher verwendete Version installieren!", MsgBoxStyle.Critical, "ITScope-Lizenz Update")

    End Sub

    Private Function GetLizenzdatum() As Date

        Dim genericConn As Sagede.OfficeLine.Data.IGenericConnection

        genericConn = Connector._mandant.MainDevice.GenericConnection
        Dim lookup As New Sagede.OfficeLine.Data.GenericDataLookup(genericConn)

        Return IDecodeString(lookup.GetString("Bezeichnung", "KHKGruppen", "Mandant=" & Connector._mandant.Id & " AND Typ=50900 AND Gruppe='License'"), "schwindt")

    End Function

    Private Function GetServerDate() As String

        Dim myWebClient As WebClient
        myWebClient = New WebClient

        Dim myResult As String = ""
        Dim dlString As String

        dlString = "http://to-officeline.de/Webservice.svc/Date"

        Try
            myResult = myWebClient.DownloadString(dlString)
            myResult = Mid(myResult, 2, Len(myResult) - 2)
        Catch ex As Exception

        End Try

        Return myResult

    End Function

    Private Function GetMaxVersion(ByVal KundenNr As String)

        Dim myWebClient As WebClient
        myWebClient = New WebClient

        Dim myResult As String = ""
        Dim dlString As String
        Dim sQry As String

        dlString = "http://to-officeline.de/Webservice.svc/KundenNr=" & KundenNr & "/Anwendung=ITS/Version"

        Try
            myResult = myWebClient.DownloadString(dlString)
            myResult = Mid(myResult, 2, Len(myResult) - 2)
        Catch ex As Exception

        End Try

        Return myResult

    End Function

    Public Function Dez2Bin(nWert As Integer) As String

        Dim i As Integer
        Dim sRest As String = ""
        Dim nErg As Integer
        Dim oCol As New Collection

        nErg = 1

        Do While nErg >= 1
            nErg = nWert \ 2
            sRest = nWert Mod 2 & sRest
            nWert = nErg
        Loop

        If Len(sRest) < 4 Then
            For i = 1 To 4 - Len(sRest)
                sRest = "0" & sRest
            Next
        End If

        Dez2Bin = sRest

    End Function

    Public Function IDecodeString(ByVal strToDecode As String, _
  ByVal strPassword As String) As String

        Dim strResult As String = ""
        Dim i As Long
        Dim cfc() As Integer
        Dim ttc() As Integer

        ReDim cfc(Len(strPassword))
        ReDim ttc(Len(strToDecode))

        For i = 1 To UBound(cfc)
            cfc(i) = Asc(Right(strPassword, _
              Len(strPassword) - i + 1))
        Next i

        For i = 1 To Len(strToDecode)
            strResult = strResult & _
              Chr(GetOfIndex(Asc(Right(strToDecode, _
              Len(strToDecode) - i + 1)), VirtPos(i, cfc)))
        Next i

        IDecodeString = strResult

    End Function

    Private Function GetOfIndex(i As Integer, _
  j As Integer) As Integer

        If i - j < 0 Then
            GetOfIndex = i - j + 255
        Else
            GetOfIndex = i - j
        End If
    End Function

    Private Function VirtPos(i As Long, _
  a() As Integer) As Integer

        If i > UBound(a) Then
            VirtPos = VirtPos(i - UBound(a), a)
        Else
            VirtPos = a(i)
        End If
    End Function

    ' Text in Verbindung mit einem Passwort verschlüsseln
    Public Function IEncodeString(ByVal strToEncode As String, _
      ByVal strPassword As String) As String

        Dim strResult As String = ""
        Dim i As Long
        Dim cfc() As Integer

        ReDim cfc(0 To Len(strPassword))

        For i = 1 To UBound(cfc)
            cfc(i) = Asc(vb6_right(strPassword, _
              Len(strPassword) - i + 1))
        Next i

        For i = 1 To Len(strToEncode)
            strResult = strResult & _
              Chr(addToIndex(Asc(vb6_right(strToEncode, _
              Len(strToEncode) - i + 1)), VirtPos(i, cfc)))
        Next i

        IEncodeString = strResult

    End Function

    ' Hilfsfunktionen
    Private Function addToIndex(ByVal i As Integer, _
      ByVal j As Integer) As Integer

        If i + j > 255 Then
            addToIndex = i + j - 255
        Else
            addToIndex = i + j
        End If
    End Function

    Private Function vb6_right(ByVal sString As String, ByVal nLaenge As Integer)

        Dim nGesamtLaenge As Integer

        Dim sNeuerString As String = ""

        Dim i As Integer

        nGesamtLaenge = sString.Length

        'Die letzten zeichen abschneiden

        For i = 1 To nLaenge

            sNeuerString = sString.Substring(nGesamtLaenge - i, 1) & sNeuerString

        Next

        Return sNeuerString

    End Function


End Module
