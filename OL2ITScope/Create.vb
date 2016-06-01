
Imports Sagede.OfficeLine.Shared
Imports Sagede.Core.Logging
Module Create

    Public Function ICreateArticleEAN(ean As String) As String

        Dim product As CProduct
        Dim article As New CArticle
        Dim Vorlageartikel As String
        Dim temp As String
        Dim result As MsgBoxResult

        product = IReadProductByEAN(ean)
        Vorlageartikel = GetVorlageartikel(product.SetId)

        Dim ArticleNew As String = GetNewArticleNr(product)

        If Vorlageartikel = "" Then
            MsgBox("Für die Produktgruppe " & product.SetName & " (" & product.SetId & ") wurde noch kein Vorlageartikel definiert!", MsgBoxStyle.Exclamation, "Vorlage fehlt")
            Return False
        End If

        'Prüfung, ob Artikel bereits angelegt
        If CheckArticleITscopeID(product.puid) > "" Then
            temp = CheckArticleITscopeID(product.puid)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            result = MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!" & vbNewLine & "Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Existiert/Verknüpfen")
            If result = MsgBoxResult.Yes Then
                SetPuid4Article(product.puid, temp)
                Return temp
            End If
            Return temp
        ElseIf CheckArticleEAN(product.ean) > "" Then
            temp = CheckArticleEAN(product.ean)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            result = MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!" & vbNewLine & "Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Existiert/Verknüpfen")
            If result = MsgBoxResult.Yes Then
                SetPuid4Article(product.puid, temp)
                Return temp
            End If
            Return temp
        ElseIf CheckArticleHArtikelnummer(product.manufacturerSKU) > "" Then
            temp = CheckArticleHArtikelnummer(product.manufacturerSKU)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            result = MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!" & vbNewLine & "Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Existiert/Verknüpfen")
            If result = MsgBoxResult.Yes Then
                SetPuid4Article(product.puid, temp)
                Return temp
            End If
            Return temp
        End If

        Try
            article.CreateArticle(ArticleNew, Vorlageartikel, product)
            Return ArticleNew

        Catch ex As Exception
            Return ""
        End Try

    End Function

    Public Function ICreateArticlePuid(puid As String, Optional Silent As Boolean = False) As CProduct

        Dim product As CProduct
        Dim article As New CArticle
        Dim Vorlageartikel As String
        Dim temp As String

        product = IReadProductByPuid(puid)
        Vorlageartikel = GetVorlageartikel(product.SetId)

        Dim ArticleNew As String = GetNewArticleNr(product)

        If Vorlageartikel = "" Then
            If Silent = False Then MsgBox("Für die Produktgruppe " & product.SetName & " (" & product.SetId & ") wurde noch kein Vorlageartikel definiert!", MsgBoxStyle.Exclamation, "Vorlage fehlt")
            product.errorMessage = "Für die Produktgruppe " & product.SetName & " (" & product.SetId & ") wurde noch kein Vorlageartikel definiert!"
            Return product
        End If

        'Prüfung, ob Artikel bereits angelegt
        If CheckArticleITscopeID(product.puid) > "" Then
            temp = CheckArticleITscopeID(product.puid)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            If Silent = False Then MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!", MsgBoxStyle.Exclamation, "Existiert bereits")
            Return product
        ElseIf CheckArticleEAN(product.ean) > "" Then
            temp = CheckArticleEAN(product.ean)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            If Silent = False Then MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!", MsgBoxStyle.Exclamation, "Existiert bereits")
            Return product
        ElseIf CheckArticleHArtikelnummer(product.manufacturerSKU) > "" Then
            temp = CheckArticleHArtikelnummer(product.manufacturerSKU)
            product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
            If Silent = False Then MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!", MsgBoxStyle.Exclamation, "Existiert bereits")
            Return product
        End If

        Try
            If article.CreateArticle(ArticleNew, Vorlageartikel, product, Silent) Then product.Article = ArticleNew
            Return product

        Catch ex As Exception
            product.errorMessage = ex.Message
            Return product
        End Try

    End Function

    Public Function IStrCreateArticlePuid(puid As String, CreateMessage As Boolean, Optional Silent As Boolean = False, Optional Shopware As Boolean = False) As String

        Dim product As CProduct
        Dim article As New CArticle
        Dim Vorlageartikel As String
        Dim temp As String
        Dim result As MsgBoxResult

        product = IReadProductByPuid(puid)
        Vorlageartikel = GetVorlageartikel(product.SetId)

        AusgabeInTextFile("Vorlageartikel: " & Vorlageartikel)

        If Vorlageartikel = "" Then
            If Silent = False Then MsgBox("Für die Produktgruppe " & product.SetName & " (" & product.SetId & ") wurde noch kein Vorlageartikel definiert!", MsgBoxStyle.Exclamation, "Vorlage fehlt")
            product.errorMessage = "Für die Produktgruppe " & product.SetName & " (" & product.SetId & ") wurde noch kein Vorlageartikel definiert!"
            Return ""
        End If

        Dim ArticleNew As String = GetNewArticleNr(product)
        AusgabeInTextFile("ArticleNew: " & ArticleNew)

        If ArticleNew = "" Then
            Dim FormAnlage As New frmArtikelAnlage

            If CheckArticleITscopeID(product.puid) > "" Then
                temp = CheckArticleITscopeID(product.puid)
                product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
                If Silent = False And Shopware = False Then MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!", MsgBoxStyle.Exclamation, "Existiert bereits")
                Return temp
            End If

            FormAnlage.lblHerstellerArtikelnr.Text = "Hersteller-Artikelnr.: " & product.manufacturerSKU
            FormAnlage.txtArtikel.Text = product.manufacturerSKU
            FormAnlage.txtTemp.Text = product.manufacturerSKU
            FormAnlage.txtPuid.Text = product.puid

            FormAnlage.ShowDialog()

            Do While FormAnlage.Visible = True
                System.Windows.Forms.Application.DoEvents()
            Loop

            If FormAnlage.txtTemp.Text = "Link" Then Return FormAnlage.txtArtikel.Text

            If FormAnlage.txtArtikel.Text = "" Or FormAnlage.txtArtikel.Text = product.manufacturerSKU Then
                Return ""
            Else
                ArticleNew = FormAnlage.txtArtikel.Text
            End If
        Else
            'Prüfung, ob Artikel bereits angelegt
            If CheckArticleITscopeID(product.puid) > "" Then
                temp = CheckArticleITscopeID(product.puid)
                product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
                If Silent = False And Shopware = False Then MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!", MsgBoxStyle.Exclamation, "Existiert bereits")
                Return temp
            ElseIf CheckArticleEAN(product.ean) > "" Then
                temp = CheckArticleEAN(product.ean)
                product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
                If Silent = False Then
                    result = MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!" & vbNewLine & "Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Existiert/Verknüpfen")
                    If result = MsgBoxResult.Yes Then
                        SetPuid4Article(product.puid, temp)
                        Return temp
                    End If
                End If

                Return temp
            ElseIf CheckArticleHArtikelnummer(product.manufacturerSKU) > "" Then
                temp = CheckArticleHArtikelnummer(product.manufacturerSKU)
                product.errorMessage = "Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!"
                If Silent = False Then
                    result = MsgBox("Der Artikel ist unter der Nummer " & temp & " bereits vorhanden!" & vbNewLine & "Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Existiert/Verknüpfen")
                    If result = MsgBoxResult.Yes Then
                        SetPuid4Article(product.puid, temp)
                        Return temp
                    End If
                End If

                Return temp
            End If
        End If

        Try
            AusgabeInTextFile("Vor Artikelanlage")
            article.CreateArticle(ArticleNew, Vorlageartikel, product, Silent, CreateMessage)
            AusgabeInTextFile(ArticleNew)

            product.Article = ArticleNew
            Return ArticleNew

        Catch ex As Exception
            Return ""
        End Try

    End Function

End Module
