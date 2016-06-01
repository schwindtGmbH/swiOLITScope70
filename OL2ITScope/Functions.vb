Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web
Imports Newtonsoft.Json
Imports Sagede.Core.Tools

Module Functions

    Dim ZeichenListe() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}

    Public Function GetLogin64Base() As String

        Dim APIKey As String = Connector._mandant.PropertyManager.GetValue(74903).ToString
        Dim Login As String = "user:" & APIKey

        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(Login)
        Dim Login64Base As String = Convert.ToBase64String(byt)

        Return Login64Base

    End Function

    Public Function GetPathForPdf() As String

        Dim Path As String = Connector._mandant.PropertyManager.GetValue(74904).ToString

        Return Path

    End Function

    Public Function GetPathForPictures() As String

        Dim Path As String = Connector._mandant.PropertyManager.GetValue(74911).ToString

        Return Path

    End Function

    Public Function FindNumbers(ByVal Source As String) As String
        If Source = "" Then Return String.Empty
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim tca As Char() = Source.ToCharArray()
        For Each Chars As Char In tca
            If Char.IsDigit(Chars) = True Then sb.Append(Chars)
        Next
        Return sb.ToString()
    End Function

    Public Function GetVorlageartikel(SetId As Int64) As String

        Dim Artikel As String

        Artikel = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Vorlageartikel", "swiITscopeArtikelgruppen", "id=" & SetId & " AND Mandant =" & Connector._mandant.Id & ""))
        Return Artikel

    End Function

    Public Function GetArtikelgruppe(SetId As Int64) As String

        Dim Artikelgruppe As String

        Artikelgruppe = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Artikelgruppe", "swiITscopeArtikelgruppen", "id=" & SetId & " AND Mandant =" & Connector._mandant.Id & ""))
        Return Artikelgruppe

    End Function

    Public Function GetSteuercode(Steuersatz As Decimal) As Int32

        Dim Steuercode As Int32

        Steuercode = Convert.ToInt32(Connector._mandant.MainDevice.Lookup.GetInt32("Steuercode", "KHKSteuertabelle", "Steuersatz=" & Steuersatz & " AND Land ='*' AND Steuertyp =0"))
        Return Steuercode

    End Function

    Public Function GetSteuerklasse(Steuercode As Int32) As Int32

        Dim Steuerklasse As Int32

        Steuerklasse = Convert.ToInt32(Connector._mandant.MainDevice.Lookup.GetInt32("Steuerklasse", "KHKSteuerklassen", "Lieferung=" & Steuercode & " AND Land ='*'"))
        Return Steuerklasse

    End Function

    Public Function GetNewArticleNr(product As CProduct) As String

        Dim Artikel As String
        Dim ArtikelMax As String
        Dim ArtikelNumerisch As Int64
        Dim ArtikelString As String
        Dim Nullen As Integer = 0
        Dim Praefix As Integer = 0
        Dim Zeichen As Integer = 0
        Dim sb As StringBuilder = New StringBuilder
        Dim a As Integer = Asc("A"c)

        Artikel = Connector._mandant.MainDevice.Lookup.Max("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & "")
        AusgabeInTextFile("MaxArtikelnummer: " & Artikel)

        If Convert.ToString(Connector._mandant.PropertyManager.GetValue(74905).ToString) = "0" Then
            ArtikelNumerisch = Convert.ToInt64(Artikel) + 1
            Nullen = Len(Artikel) - Len(ArtikelNumerisch)
            Artikel = ArtikelNumerisch
            For i = 1 To Nullen
                Artikel = "0" & Artikel
            Next
        ElseIf Convert.ToString(Connector._mandant.PropertyManager.GetValue(74905).ToString) = "-1" Then
            If Convert.ToString(Connector._mandant.PropertyManager.GetValue(74912).ToString) <> "0" Then
                Artikel = Convert.ToString(Connector._mandant.PropertyManager.GetValue(74912).ToString)
                AusgabeInTextFile("Alphanumerisch: " & Artikel)
                'ArtikelMax = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer LIKE '" & Artikel & "%'" & " AND LEN(Artikelnummer)=8 ORDER BY Artikelnummer DESC"))

                Praefix = Len(Convert.ToString(Connector._mandant.PropertyManager.GetValue(74912)))
                AusgabeInTextFile("Praefix: " & Praefix)
                Nullen = Convert.ToInt16(Connector._mandant.PropertyManager.GetValue(74913))
                AusgabeInTextFile("Nullen: " & Nullen)
                ArtikelMax = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer LIKE '" & Artikel & "%'" & " AND LEN(Artikelnummer)=" & Nullen + Praefix & " ORDER BY Artikelnummer DESC"))
                AusgabeInTextFile("ArtikelMax: " & ArtikelMax)
                If ArtikelMax = "" Then
                    Nullen = Convert.ToInt16(Connector._mandant.PropertyManager.GetValue(74913))
                    For i = 1 To Nullen
                        Artikel = Artikel & "0"
                    Next
                Else
                    ArtikelNumerisch = Convert.ToInt64(FindNumbers(ArtikelMax))
                    ArtikelNumerisch = ArtikelNumerisch + 1
                    ArtikelString = Convert.ToString(ArtikelNumerisch)
                    Zeichen = Len(ArtikelMax) - Len(ArtikelString)
                    Artikel = Convert.ToString(Left(Artikel, Zeichen) & ArtikelNumerisch)
                End If
            Else
                Artikel = ""
            End If
            AusgabeInTextFile("Artikel: " & Artikel)
        Else
            'Hersteller-Artikelnr.
            If CheckArticleHArtikelnummer(product.manufacturerSKU) = "" Then
                Artikel = product.manufacturerSKU
            Else
                'Hersteller-Artikelnr. bereits belegt
                Artikel = ""
            End If
        End If

        Return Artikel

    End Function

    Public Sub SetPuid4Article(puid As String, article As String)

        Dim qry As String = "UPDATE KHKArtikel SET USER_swiITscopeNummer ='" & puid & "' WHERE Mandant =" & Connector._mandant.Id & " AND Artikelnummer ='" & article & "'"
        Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(qry)

        MsgBox("Der Artikel wurde erfolgreich mit ITscope-ID " & puid & " verknüpft!", MsgBoxStyle.Information, "Erfolgreich")

    End Sub

    Public Sub ICheckPuid4Article(article As String)

        Dim puid As String
        Dim ean As String
        Dim result As MsgBoxResult

        'Zuerst checken, ob schon zugeordnet
        puid = ConversionHelper.ToString(Connector._mandant.MainDevice.Lookup.GetString("USER_swiITscopeNummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer='" & article & "' And Aktiv = -1"))
        If puid > "" Then
            MsgBox("Der Artikel ist bereits einem ITscope-Produkt zugeordnet!", vbExclamation, "Bereits zugeordnet!")
            Exit Sub
        Else
            result = MsgBox("Wollen Sie den Artikel: " & article & " mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Verknüpfen")
            If result = MsgBoxResult.Yes Then

                ean = ConversionHelper.ToString(Connector._mandant.MainDevice.Lookup.GetString("EANNummer", "KHKArtikelVarianten", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer='" & article & "' AND AuspraegungID = 0"))

                If ean <> "" Then
                    puid = GetPuid4EAN(ean)
                    If puid <> "" Then
                        SetPuid4Article(puid, article)
                    Else
                        MsgBox("Der Artikel konnte nicht in ITscope zugeordnet werden!", vbExclamation, "Keine Zuordnung!")
                    End If
                Else
                    MsgBox("Der Artikel hat keine EAN für eine Verknüpfung hinterlegt!", vbExclamation, "Keine EAN!")
                End If
            End If
        End If

    End Sub

    Public Function GetPuid4EAN(ean As String) As String

        Dim Request As HttpWebRequest = CType(WebRequest.Create("https://api.itscope.com/1.0/products/ean/" & ean & "/full.json?realtime=true"), HttpWebRequest)
        Dim Login64Base As String = GetLogin64Base()

        Request.Method = "GET"
        Request.ContentType = "application/json"

        Request.Headers.Add(HttpRequestHeader.Authorization, Login64Base)
        Request.Headers.Add(HttpRequestHeader.AcceptLanguage, "de")

        Try
            Dim Response As HttpWebResponse = Request.GetResponse()
            Dim DataStream As Stream

            DataStream = Response.GetResponseStream()
            Dim reader As New StreamReader(DataStream)
            Dim ServerResponse As String = reader.ReadToEnd()

            Dim dict As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of IDictionary(Of String, Object))(ServerResponse)

            Return dict.Item("product")(0)("puid").ToString

        Catch ex As Exception
            Return ""
        End Try

    End Function

    Public Function CheckArticleITscopeID(id As String) As String

        Dim Artikel As String

        Artikel = Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND USER_swiITscopeNummer='" & id & "' And Aktiv = -1")

        If Artikel = "" Then
            Return Artikel
        Else
            Return Artikel
        End If

    End Function

    Public Function CheckArticleEAN(ean As String) As String

        Dim Artikel As String

        Artikel = Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikelVarianten", "Mandant = " & Connector._mandant.Id & " AND EANNummer='" & ean & "'")
        Artikel = Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer='" & Artikel & "' AND Aktiv=-1")

        If Artikel = "" Then
            Return Artikel
        Else
            Return Artikel
        End If

    End Function

    Public Function CheckArticleHArtikelnummer(HArtikelnummer As String) As String

        Dim Artikel As String

        Artikel = Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND HArtikelnummer ='" & HArtikelnummer & "' AND Aktiv = -1")

        If Artikel = "" Then
            Return Artikel
        Else
            Return Artikel
        End If

    End Function

    Public Function CheckArtikelnummer(Artikelnummer As String) As Boolean

        Dim Artikel As String

        Artikel = Connector._mandant.MainDevice.Lookup.GetString("Artikelnummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer ='" & Artikelnummer & "'")

        If Artikel = "" Then
            Return False
        Else
            Return True
        End If

    End Function

    Function Count(ByVal Text As String) As String

        Dim ErstesChar As String                        ' 1. Zeichen in ZeichenListe
        Dim LetztesChar As String                       ' Letztes Zeichen in ZeichenListe

        Dim ÜbertragAufLinkeStelle As Boolean
        Dim CounterArray() As Char
        Dim Zeichen As String
        Dim PositionInCounterArray As Integer
        Dim PositionInCharListe As Integer

        ErstesChar = ZeichenListe.First
        LetztesChar = ZeichenListe.Last

        CounterArray = Text.ToArray                     ' 1 Stelle mehr, da eventuell Übertrag vor die 1. Stelle anfällt

        If Text = String.Empty Then                     ' Falls Text leer ist, dann ...
            Return ErstesChar                           ' ... erstes Zeichen aus CharListe nehmen
        End If

        ÜbertragAufLinkeStelle = False

        For PositionInCounterArray = CounterArray.Length - 1 To 0 Step -1    ' Von hinten nach vorn durchgehen
            Zeichen = CounterArray(PositionInCounterArray)   ' Welches Zeichen steht dort

            'Debug.Write("PositionInCounterArray = " & PositionInCounterArray) ' Für Testzwecke
            'Debug.Write("   Zeichen = " & Zeichen)
            'Debug.WriteLine("   Übertrag = " & ÜbertragAufLinkeStelle)

            If Zeichen = LetztesChar Then
                CounterArray(PositionInCounterArray) = CChar(ErstesChar)   ' Ersetzten durch 1. Zeichen aus ZeichenListe
                ÜbertragAufLinkeStelle = True
            Else
                PositionInCharListe = Array.IndexOf(ZeichenListe, Zeichen) ' Position des Zeichens in ZeichenListe
                CounterArray(PositionInCounterArray) = CChar(ZeichenListe(PositionInCharListe + 1)) ' Nächstes Zeichen aus ZeichenListe einsetzen
                ÜbertragAufLinkeStelle = False          ' Schleife verlassen weil kein Übertrag notwendig
                Exit For
            End If
        Next

        If ÜbertragAufLinkeStelle = False Then          ' Es ist kein Übertrag auf neue 1. Stelle erforderlich
            Return String.Concat(CounterArray)
        Else                                            ' Übertrag auf neue 1. Stelle erforderlich
            Return ErstesChar & String.Concat(CounterArray)
        End If

    End Function

    Public Function StripTags(html As String) As String
        Return Regex.Replace(html, "<.*?>", "")
    End Function

End Module
