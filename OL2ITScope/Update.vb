Imports Sagede.OfficeLine.Engine
Imports Sagede.OfficeLine.Shared.Customizing
Imports Sagede.OfficeLine.Data
Imports Sagede.Core.Tools
Imports Sagede.OfficeLine.Data.Entities.Main
Imports Sagede.OfficeLine.Data.Entities

Module Update


    Public Function IUpdateArticlePuid(artikel As String, puid As String) As Boolean

        Dim _BeschaffungTemp As New ArtikelLieferantItem
        Dim _Beschaffung As New ArtikelLieferantItem
        Dim product As CProduct
        Dim article As New CArticle
        Dim kto As String
        Dim _adapter As IEntityAdapter
        Dim Qry As String
        Dim command As IGenericCommand
        Dim BME As String = ""

        product = IReadPricesByPuid(puid)

        Try
            'Bezugsquellen leeren
            Qry = "DELETE FROM swiITscopeBezugsquellen WHERE Artikel ='" & artikel & "'"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(Qry)

            For Each supplier In product.supplieritem
                kto = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Kto", "swiITscopeDistributoren", "id =" & supplier.id & " AND Mandant =" & Connector._mandant.Id & ""))
                If kto = "" Then Continue For

                Qry = "INSERT INTO swiITscopeBezugsquellen (Artikel,Mandant,Lieferant,Datum,Lagerbestand,Lagertext,Preis, priceSourceId) VALUES (@Artikel,@Mandant,@Lieferant,@Datum,@Lagerbestand,@Lagertext,@Preis,@priceSourceId)"
                command = Connector._mandant.MainDevice.GenericConnection.CreateSqlStringCommand(Qry)

                command.AppendInParameter("Artikel", GetType(String), artikel)
                command.AppendInParameter("Mandant", GetType(Int16), Connector._mandant.Id)
                command.AppendInParameter("Lieferant", GetType(String), kto)
                command.AppendInParameter("Datum", GetType(DateTime), Now)
                command.AppendInParameter("Lagerbestand", GetType(Int32), supplier.stock)
                command.AppendInParameter("Lagertext", GetType(String), supplier.stockStatusText)
                command.AppendInParameter("Preis", GetType(Double), supplier.price)
                command.AppendInParameter("priceSourceId", GetType(Integer), supplier.priceSourceId)

                command.ExecuteNonQuery()

                '-----------------------------------------------------------------------------------------------------------------------------------------------
                Qry = "INSERT INTO KHKArtikelLieferantPreise (Lieferant,Artikelnummer,Mandant,AuspraegungID,AbMenge,Einzelpreis) VALUES (@Lieferant,@Artikelnummer,@Mandant,@AuspraegungID,@AbMenge,@Einzelpreis)"
                command = Connector._mandant.MainDevice.GenericConnection.CreateSqlStringCommand(Qry)

                command.AppendInParameter("Lieferant", GetType(String), kto)
                command.AppendInParameter("Artikelnummer", GetType(String), artikel)
                command.AppendInParameter("Mandant", GetType(Int16), Connector._mandant.Id)
                command.AppendInParameter("AuspraegungID", GetType(Int16), 0)
                command.AppendInParameter("AbMenge", GetType(Int16), 1)
                command.AppendInParameter("Einzelpreis", GetType(Double), supplier.price)

                _adapter = Connector._mandant.MainDevice.Entities.Adapter

                _BeschaffungTemp = Connector._mandant.MainDevice.Entities.ArtikelLieferant.GetItem(kto, artikel, Connector._mandant.Id, 0, False)
                If Not _BeschaffungTemp Is Nothing Then
                    _BeschaffungTemp.Einzelpreis = supplier.price

                    _BeschaffungTemp.Save()
                Else
                    _BeschaffungTemp = _BeschaffungTemp.Create(_adapter, kto, artikel, Connector._mandant.Id, 0)

                    BME = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Basismengeneinheit", "KHKArtikel", "Mandant =" & Connector._mandant.Id & " AND Artikelnummer ='" & artikel & "'"))
                    _BeschaffungTemp.Artikelnummer = artikel
                    _BeschaffungTemp.AuspraegungID = 0
                    _BeschaffungTemp.Bestellnummer = supplier.supplierItemId
                    _BeschaffungTemp.Bezeichnung1 = ""
                    _BeschaffungTemp.Bezeichnung2 = ""
                    _BeschaffungTemp.DezimalstellenEK = 0
                    _BeschaffungTemp.Einkaufsmengeneinheit = BME
                    _BeschaffungTemp.Einzelpreis = supplier.price
                    _BeschaffungTemp.MengenberechnungEK = Nothing
                    _BeschaffungTemp.Mindestbestellmenge = 0
                    _BeschaffungTemp.PreiseinheitEK = 1
                    _BeschaffungTemp.Rabattsatz = 0
                    _BeschaffungTemp.UmrechnungsfaktorEK = 1
                    _BeschaffungTemp.UmrechnungsfaktorVPEK = 1
                    _BeschaffungTemp.Wiederbeschaffungszeit = 0

                    _BeschaffungTemp.Save()
                End If
            Next

            Return True

        Catch ex As Exception
            product.errorMessage = ex.Message
            Return False
        End Try

    End Function

    Public Function IUpdateArticleFull(article As String) As Boolean

        Dim puid As String
        Dim product As CProduct
        Dim _Artikel As New ArtikelItem
        Dim _adapter As IEntityAdapter
        Dim i, y As Integer
        Dim result As MsgBoxResult

        'Zuerst checken, ob schon zugeordnet
        puid = ConversionHelper.ToString(Connector._mandant.MainDevice.Lookup.GetString("USER_swiITscopeNummer", "KHKArtikel", "Mandant = " & Connector._mandant.Id & " AND Artikelnummer='" & article & "' And Aktiv = -1"))
        If puid = "" Then
            MsgBox("Der Artikel ist keinem einem ITscope-Produkt zugeordnet!", vbExclamation, "Keine Zuordnung!")
            Return False
        End If

        result = MsgBox("Wollen Sie den Artikel: " & article & " aktualisieren?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Aktualisieren")
        If result = MsgBoxResult.No Then
            Return False
        End If

        'Produktdaten einlesen
        product = IReadProductByPuid(puid)

        Try
            _adapter = Connector._mandant.MainDevice.Entities.Adapter
            _Artikel = Connector._mandant.MainDevice.Entities.Artikel.GetItem(article, Connector._mandant.Id)

            'Prüfe, ob Hersteller in Bezeichnung angezeigt werden soll, 02.03.2016, MG, Umstellung API 2.0, Funktion deaktiviert, da Hersteller nun immer automatisch im Produktnamen
            'If ConversionHelper.ToInt32(Connector._mandant.PropertyManager.GetValue(74915).ToString) = -1 Then product.productName = product.manufacturer & " " & product.productName

            i = InStr(35, product.productName, " ", vbTextCompare)

            If i = 0 Then
                _Artikel.Bezeichnung1 = product.productName
                If Len(_Artikel.Bezeichnung1) > 50 Then _Artikel.Bezeichnung1 = Left(product.productName, 50)
            Else
                _Artikel.Bezeichnung1 = Left(product.productName, i)
            End If

            If i <> 0 Then
                y = InStr(i + 35, product.productName, " ", vbTextCompare)
                If y = 0 Then
                    _Artikel.Bezeichnung2 = Mid(product.productName, i, (Len(product.productName) - i + 1))
                Else
                    _Artikel.Bezeichnung2 = Mid(product.productName, i, y - i + 1)
                End If
                If Len(_Artikel.Bezeichnung2) > 50 Then
                    _Artikel.Bezeichnung2 = Left(_Artikel.Bezeichnung2, 50)
                End If
            End If

            _Artikel.Matchcode = _Artikel.Bezeichnung1
            _Artikel.Bezeichnung2 = LTrim(_Artikel.Bezeichnung2)

            Dim Mode As Integer

            Mode = ConversionHelper.ToInt32(Connector._mandant.PropertyManager.GetValue(74907).ToString)
            Select Case Mode

                Case -1
                    _Artikel.Dimensionstext = ""
                Case 0, 1
                    _Artikel.Dimensionstext = product.Dimensionstext
                Case Else
                    _Artikel.Dimensionstext = StripTags(ConversionHelper.ToString(product.Dimensionstext))
                    _Artikel.DimensionstextHTML = product.Dimensionstext
                    _Artikel.DimensionstextHTML = Replace(product.Dimensionstext, "\\n", "<br>")
                    'MG, 10.02.2016, Änderung RTF-Convert auf Basis des DimensionstextHTML
                    '_Artikel.DimensionstextRTF = ConversionHelper.HtmlToRtf(product.Dimensionstext)
                    _Artikel.DimensionstextRTF = ConversionHelper.HtmlToRtf(_Artikel.DimensionstextHTML)
            End Select

            Mode = ConversionHelper.ToInt32(Connector._mandant.PropertyManager.GetValue(74908).ToString)
            Select Case Mode

                Case -1
                    _Artikel.Langtext = ""
                Case 0, 1
                    _Artikel.Langtext = product.Langtext
                Case Else
                    _Artikel.Langtext = StripTags(ConversionHelper.ToString(product.Langtext))
                    _Artikel.LangtextHTML = product.Langtext
                    _Artikel.LangtextHTML = Replace(_Artikel.LangtextHTML, "\\n", "<br>", 1, -1, CompareMethod.Text)
                    'MG, 10.02.2016, Änderung RTF-Convert auf Basis des LangtexttHTML
                    '_Artikel.LangtextRTF = ConversionHelper.HtmlToRtf(product.Langtext)
                    _Artikel.LangtextRTF = ConversionHelper.HtmlToRtf(_Artikel.LangtextHTML)
            End Select

            If Len(product.manufacturer) > 20 Then
                i = InStr(1, product.manufacturer, " ", vbTextCompare)
                If i > 0 Then
                    _Artikel.Hersteller = Left(product.manufacturer, i - 1)
                    If Len(_Artikel.Hersteller) > 20 Then Left(product.manufacturer, 20)
                Else
                    _Artikel.Hersteller = Left(product.manufacturer, 20)
                End If
            Else
                _Artikel.Hersteller = product.manufacturer
            End If

            _Artikel.Eigenmasse = product.estimateGrossWeight
            _Artikel.HArtikelnummer = Left(product.manufacturerSKU, 20)
            _Artikel.Breite = ConversionHelper.ToDecimal(FindNumbers(product.grossDimX)) / 10
            _Artikel.Hoehe = ConversionHelper.ToDecimal(FindNumbers(product.grossDimY)) / 10
            _Artikel.Laenge = ConversionHelper.ToDecimal(FindNumbers(product.grossDimZ)) / 10

            _Artikel.Save()

            'Alle Bilder in Sammelmappe löschen
            Dim qry As String = ""
            qry = "DELETE FROM KHKSammelmappen WHERE Schluessel ='" & _Artikel.Artikelnummer & "' AND Mandant =" & Connector._mandant.Id & " AND Typ = 2 AND Mappe=-7"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(qry)

            'Pictures in Sammelmappe
            If Connector._mandant.PropertyManager.GetValue(74910).ToString = "-1" Then
                Dim picture As CPictures
                Dim Pfad As String = GetPathForPictures()
                Dim ArticleFunctions As New CArticle

                If Pfad = "" Or Pfad = "0" Then
                    MsgBox("Es wurde noch kein Bilder-Pfad in den Grundlagen hinterlegt!", MsgBoxStyle.Exclamation, "Bild-Pfad")
                Else
                    For Each picture In product.pictureitem
                        ArticleFunctions.FillPicture(article, picture.name, picture.value, False)
                    Next
                End If
            End If

            'Bezugsquellen leeren
            Dim kto As String
            Dim command As IGenericCommand
            Dim _BeschaffungTemp As New ArtikelLieferantItem
            Dim BME As String = ""

            qry = "DELETE FROM swiITscopeBezugsquellen WHERE Artikel ='" & article & "'"
            Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(qry)

            For Each supplier In product.supplieritem
                kto = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Kto", "swiITscopeDistributoren", "id =" & supplier.id & " AND Mandant =" & Connector._mandant.Id & ""))
                If kto = "" Then Continue For

                qry = "INSERT INTO swiITscopeBezugsquellen (Artikel,Mandant,Lieferant,Datum,Lagerbestand,Lagertext,Preis, priceSourceId) VALUES (@Artikel,@Mandant,@Lieferant,@Datum,@Lagerbestand,@Lagertext,@Preis,@priceSourceId)"
                command = Connector._mandant.MainDevice.GenericConnection.CreateSqlStringCommand(qry)

                command.AppendInParameter("Artikel", GetType(String), article)
                command.AppendInParameter("Mandant", GetType(Int16), Connector._mandant.Id)
                command.AppendInParameter("Lieferant", GetType(String), kto)
                command.AppendInParameter("Datum", GetType(DateTime), Now)
                command.AppendInParameter("Lagerbestand", GetType(Int32), supplier.stock)
                command.AppendInParameter("Lagertext", GetType(String), supplier.stockStatusText)
                command.AppendInParameter("Preis", GetType(Double), supplier.price)
                command.AppendInParameter("priceSourceId", GetType(Integer), supplier.priceSourceId)

                command.ExecuteNonQuery()

                '-----------------------------------------------------------------------------------------------------------------------------------------------
                qry = "INSERT INTO KHKArtikelLieferantPreise (Lieferant,Artikelnummer,Mandant,AuspraegungID,AbMenge,Einzelpreis) VALUES (@Lieferant,@Artikelnummer,@Mandant,@AuspraegungID,@AbMenge,@Einzelpreis)"
                command = Connector._mandant.MainDevice.GenericConnection.CreateSqlStringCommand(qry)

                command.AppendInParameter("Lieferant", GetType(String), kto)
                command.AppendInParameter("Artikelnummer", GetType(String), article)
                command.AppendInParameter("Mandant", GetType(Int16), Connector._mandant.Id)
                command.AppendInParameter("AuspraegungID", GetType(Int16), 0)
                command.AppendInParameter("AbMenge", GetType(Int16), 1)
                command.AppendInParameter("Einzelpreis", GetType(Double), supplier.price)

                _adapter = Connector._mandant.MainDevice.Entities.Adapter

                _BeschaffungTemp = Connector._mandant.MainDevice.Entities.ArtikelLieferant.GetItem(kto, article, Connector._mandant.Id, 0, False)
                If Not _BeschaffungTemp Is Nothing Then
                    _BeschaffungTemp.Einzelpreis = supplier.price

                    _BeschaffungTemp.Save()
                Else
                    _BeschaffungTemp = _BeschaffungTemp.Create(_adapter, kto, article, Connector._mandant.Id, 0)

                    BME = Convert.ToString(Connector._mandant.MainDevice.Lookup.GetString("Basismengeneinheit", "KHKArtikel", "Mandant =" & Connector._mandant.Id & " AND Artikelnummer ='" & article & "'"))
                    _BeschaffungTemp.Artikelnummer = article
                    _BeschaffungTemp.AuspraegungID = 0
                    _BeschaffungTemp.Bestellnummer = supplier.supplierItemId
                    _BeschaffungTemp.Bezeichnung1 = ""
                    _BeschaffungTemp.Bezeichnung2 = ""
                    _BeschaffungTemp.DezimalstellenEK = 0
                    _BeschaffungTemp.Einkaufsmengeneinheit = BME
                    _BeschaffungTemp.Einzelpreis = supplier.price
                    _BeschaffungTemp.MengenberechnungEK = Nothing
                    _BeschaffungTemp.Mindestbestellmenge = 0
                    _BeschaffungTemp.PreiseinheitEK = 1
                    _BeschaffungTemp.Rabattsatz = 0
                    _BeschaffungTemp.UmrechnungsfaktorEK = 1
                    _BeschaffungTemp.UmrechnungsfaktorVPEK = 1
                    _BeschaffungTemp.Wiederbeschaffungszeit = 0

                    _BeschaffungTemp.Save()
                End If
            Next

            MsgBox("Der Artikel wurde erfolgreich aktualisiert!", vbInformation, "Update")

            Return True

        Catch ex As Exception
            product.errorMessage = ex.Message
            MsgBox(ex.Message)
        End Try

    End Function

End Module
