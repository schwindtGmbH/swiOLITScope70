Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Sagede.OfficeLine.Data
Imports Sagede.OfficeLine.Engine
Imports Sagede.Core.Tools


Module ProductQuery

    Public Function IReadProductByEAN(ean As String) As CProduct

        Dim productEAN As CProduct

        Try
            productEAN = IReadProductByPuid("EAN-Variante", ean)

            Return productEAN

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
            Return Nothing
        End Try

    End Function

    Public Sub IReadProductDatasheetByEAN(ean As String, Optional IsMail As Boolean = False)

        AusgabeInTextFile("DatesheetByEAN")

        Dim Request As HttpWebRequest = CType(WebRequest.Create("https://api.itscope.com/2.0/products/ean/" & ean & "/standard.json"), HttpWebRequest)
        AusgabeInTextFile("EAN: " & ean)
        Dim Login64Base As String = GetLogin64Base()
        Dim Pfad As String = GetPathForPdf()
        Dim pdfPfad As String

        AusgabeInTextFile("PFD-Pfad: " & Pfad)

        If Pfad = "" Or Pfad = "0" Then
            MsgBox("Es wurde noch kein PDF-Download Pfad in den Grundlagen hinterlegt!", MsgBoxStyle.Exclamation, "PDF Pfad")
            Exit Sub
        End If

        Request.Method = "GET"
        Request.ContentType = "application/json"

        Try
            Request.Headers.Add(HttpRequestHeader.Authorization, Login64Base)
            Request.Headers.Add(HttpRequestHeader.AcceptLanguage, "de")

            Request.UserAgent = "schwindt-sage-" & GetVersion()

            Dim Response As HttpWebResponse = Request.GetResponse()
            Dim DataStream As Stream

            DataStream = Response.GetResponseStream()

            Dim reader As New StreamReader(DataStream)
            Dim ServerResponse As String = reader.ReadToEnd()

            Dim dict As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of IDictionary(Of String, Object))(ServerResponse)

            pdfPfad = dict.Item("product")(0)("standardPdfDatasheet").ToString

            Dim Request2 As HttpWebRequest = CType(WebRequest.Create(pdfPfad), HttpWebRequest)
            Request2.Method = "GET"
            Request2.ContentType = "application/x-www-form-urlencoded"

            Request2.Headers.Add(HttpRequestHeader.Authorization, Login64Base)
            Request2.Headers.Add(HttpRequestHeader.AcceptLanguage, "de")

            Request2.UserAgent = "schwindt-sage-7.0.8"

            Dim Response2 As HttpWebResponse = Request2.GetResponse()
            Dim DataStream2 As Stream

            DataStream2 = Response2.GetResponseStream()

            Using fs As New FileStream(Pfad & ean & ".pdf", FileMode.Create)
                Dim read As Byte() = New Byte(255) {}
                Dim count As Integer = DataStream2.Read(read, 0, read.Length)
                While count > 0
                    fs.Write(read, 0, count)
                    count = DataStream2.Read(read, 0, read.Length)
                End While
                AusgabeInTextFile("VorProcessStart")
                If Not IsMail Then Process.Start(Pfad & ean & ".pdf")
            End Using

            DataStream.Close()
            Response.Close()

            DataStream2.Close()
            Response2.Close()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
        End Try

    End Sub

    Public Sub IMailProductDatasheetByEAN(ean As String)

        Dim Pfad As String = GetPathForPdf()

        If Pfad = "" Or Pfad = "0" Then
            MsgBox("Es wurde noch kein PDF-Download Pfad in den Grundlagen hinterlegt!", MsgBoxStyle.Exclamation, "PDF Pfad")
            Exit Sub
        End If

        Try
            IReadProductDatasheetByEAN(ean, True)

            Dim mail As New Schwindt.Mail.MailFunctions
            mail.emailsenden("", "Datenblatt", "", Pfad & ean & ".pdf")

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
        End Try

    End Sub

    Public Function IReadProductByPuid(puid As String, Optional EAN As String = "") As CProduct

        Dim Request As HttpWebRequest

        If EAN <> "" Then
            Request = CType(WebRequest.Create("https://api.itscope.com/2.0/products/ean/" & EAN & "/standard.json?realtime=true"), HttpWebRequest)
        Else
            Request = CType(WebRequest.Create("https://api.itscope.com/2.0/products/id/" & puid & "/standard.json?realtime=true"), HttpWebRequest)
        End If

        Dim Login64Base As String = GetLogin64Base()
        Dim Mode As Integer

        Request.Method = "GET"
        Request.ContentType = "application/json"

        Request.Headers.Add(HttpRequestHeader.Authorization, Login64Base)
        Request.Headers.Add(HttpRequestHeader.AcceptLanguage, "de")

        Request.UserAgent = "schwindt-sage-" & GetVersion()

        Try

            Dim Response As HttpWebResponse = Request.GetResponse()
            Dim DataStream As Stream

            DataStream = Response.GetResponseStream()
            Dim reader As New StreamReader(DataStream)
            Dim ServerResponse As String = reader.ReadToEnd()

            Dim dict As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of IDictionary(Of String, Object))(ServerResponse)
            Dim product As New CProduct
            Dim supplier As New CSupplier
            Dim picture As New CPictures

            product.puid = dict.Item("product")(0)("puid").ToString
            product.SetId = dict.Item("product")(0)("productTypeId").ToString
            product.SetName = dict.Item("product")(0)("productTypeName").ToString

            Try
                product.manufacturer = ConversionHelper.ToString(dict.Item("product")(0)("manufacturerName").ToString)
            Catch ex As Exception

            End Try

            product.productName = ConversionHelper.ToString(dict.Item("product")(0)("productNameWithManufacturer").ToString)

            Try
                product.shortInfo = ConversionHelper.ToString(dict.Item("product")(0)("shortDescription").ToString)
            Catch ex As Exception

            End Try

            Try
                product.ean = ConversionHelper.ToString(dict.Item("product")(0)("ean").ToString)
            Catch ex As Exception

            End Try

            Try
                product.manufacturerSKU = ConversionHelper.ToString(dict.Item("product")(0)("manufacturerSKU").ToString)
            Catch ex As Exception

            End Try

            product.deeplink = ConversionHelper.ToString(dict.Item("product")(0)("deeplink").ToString)

            Try
                product.estimateGrossWeight = dict.Item("product")(0)("estimateGrossWeight").ToString
            Catch ex As Exception

            End Try

            Try
                product.vat = ConversionHelper.ToString(dict.Item("product")(0)("priceCalcVat").ToString)
            Catch ex As Exception

            End Try

            Try
                product.grossDimX = ConversionHelper.ToString(dict.Item("product")(0)("grossDimX").ToString)
                product.grossDimY = ConversionHelper.ToString(dict.Item("product")(0)("grossDimY").ToString)
                product.grossDimZ = ConversionHelper.ToString(dict.Item("product")(0)("grossDimZ").ToString)
            Catch ex As Exception

            End Try

            Try
                product.calcPrice = ConversionHelper.ToDecimal(Replace(dict.Item("product")(0)("priceCalc").ToString, ".", ","))

            Catch ex As Exception
                product.calcPrice = 0
            End Try

            If Connector._mandant.PropertyManager.GetValue(74910).ToString = "-1" Then
                For g = 1 To 5
                    Try
                        If dict.Item("product")(0)("image" & g).ToString <> "" Then
                            picture = New CPictures
                            product.pictureitem.Add(picture)
                            picture.name = dict.Item("product")(0)("image" & g).ToString
                            picture.value = dict.Item("product")(0)("image" & g).ToString
                        End If
                    Catch ex As Exception

                    End Try
                Next

            End If

            Mode = ConversionHelper.ToInt32(Connector._mandant.PropertyManager.GetValue(74907).ToString)

            Select Case Mode

                Case 0

                    product.Dimensionstext = product.shortInfo

                    If product.Dimensionstext = "" Then
                        product.Dimensionstext = product.productName
                    End If

                Case 1

                    Try
                        product.Dimensionstext = ConversionHelper.ToString(dict.Item("product")(0)("longDescription").ToString)
                    Catch ex As Exception

                    End Try

                    If product.Dimensionstext = "" Then
                        product.Dimensionstext = product.productName & vbNewLine & product.shortInfo
                    End If

                Case 2
Case2Dim:
                    Try
                        product.Dimensionstext = ConversionHelper.ToString(dict.Item("product")(0)("marketingText").ToString)
                    Catch ex As Exception

                    End Try

                Case 3

                    Try
                        product.Dimensionstext = ConversionHelper.ToString(dict.Item("product")(0)("longDescription").ToString)
                    Catch ex As Exception

                    End Try

                    If product.Dimensionstext = "" Then GoTo Case2Dim

            End Select

            'Langtext
            Mode = ConversionHelper.ToInt32(Connector._mandant.PropertyManager.GetValue(74908).ToString)

            Select Case Mode

                Case 0

                    product.Langtext = product.shortInfo

                    If product.Langtext = "" Then
                        product.Langtext = product.productName
                    End If

                Case 1

                    Try
                        product.Langtext = ConversionHelper.ToString(dict.Item("product")(0)("longDescription").ToString)
                    Catch ex As Exception

                    End Try

                    If product.Langtext = "" Then
                        product.Langtext = product.productName & vbNewLine & product.shortInfo
                    End If

                Case 2
Case2Lang:
                    Try
                        product.Langtext = ConversionHelper.ToString(dict.Item("product")(0)("marketingText").ToString)
                    Catch ex As Exception

                    End Try


                Case 3
                    Try
                        product.Langtext = ConversionHelper.ToString(dict.Item("product")(0)("longDescription").ToString)
                    Catch ex As Exception

                    End Try

                    If product.Langtext = "" Then GoTo Case2Lang

            End Select

            Dim Qry As String
            Dim rs As IGenericRecordset

            Qry = "SELECT id, name, kto FROM swiITscopeDistributoren where Mandant =" & Connector._mandant.Id & " AND kto >''"
            rs = New GenericRecordset(Connector._mandant.MainDevice)
            rs.Open(Qry, False)

            If Not rs.EOF Then

                For i = 0 To dict.Item("product")(0)("supplierItems").Count - 1
                    'Nur hinterlegte Distris ansprechen
                    Do Until rs.EOF
                        If rs.Item("id") = dict.Item("product")(0)("supplierItems")(i)("supplierId").ToString Then

                            If dict.Item("product")(0)("supplierItems")(i)("conditionId").ToString = "1" Then
                                supplier = New CSupplier
                                product.supplieritem.Add(supplier)
                                supplier.id = dict.Item("product")(0)("supplierItems")(i)("supplierId").ToString

                                Try
                                    supplier.supplierItemId = dict.Item("product")(0)("supplierItems")(i)("supplierSKU").ToString

                                Catch ex As Exception
                                End Try

                                Try
                                    supplier.stockStatusText = dict.Item("product")(0)("supplierItems")(i)("stockStatusText").ToString
                                Catch ex As Exception
                                End Try

                                supplier.stock = dict.Item("product")(0)("supplierItems")(i)("stock").ToString 
                                supplier.price = Replace(dict.Item("product")(0)("supplierItems")(i)("price").ToString, ".", ",")
                                supplier.priceSourceId = 0
                                supplier.calcprice = ConversionHelper.ToDecimal(dict.Item("product")(0)("supplierItems")(i)("priceCalc").ToString)
                            End If
                        End If
                        rs.MoveNext()
                    Loop
                    rs.MoveFirst()
                Next

            End If

            rs.Close()

            reader.Close()
            DataStream.Close()
            Response.Close()

            Return product

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
            Return Nothing
        End Try

    End Function

    Public Function IReadPricesByPuid(puid As String) As CProduct

        Dim Request As HttpWebRequest = CType(WebRequest.Create("https://api.itscope.com/2.0/products/id/" & puid & "/standardUpdate.json?realtime=true"), HttpWebRequest)
        Dim Login64Base As String = GetLogin64Base()

        Request.Method = "GET"
        Request.ContentType = "application/json"

        Request.Headers.Add(HttpRequestHeader.Authorization, Login64Base)
        Request.Headers.Add(HttpRequestHeader.AcceptLanguage, "de")

        Request.UserAgent = "schwindt-sage-" & GetVersion()

        Try

            Dim Response As HttpWebResponse = Request.GetResponse()
            Dim DataStream As Stream

            DataStream = Response.GetResponseStream()
            Dim reader As New StreamReader(DataStream)
            Dim ServerResponse As String = reader.ReadToEnd()

            Dim dict As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of IDictionary(Of String, Object))(ServerResponse)
            Dim product As New CProduct
            Dim supplier As New CSupplier

            product.puid = dict.Item("product")(0)("puid").ToString

            Dim Qry As String
            Dim rs As IGenericRecordset

            Qry = "SELECT id, name, kto FROM swiITscopeDistributoren where Mandant =" & Connector._mandant.Id & " AND kto >''"
            rs = New GenericRecordset(Connector._mandant.MainDevice)
            rs.Open(Qry, False)

            If Not rs.EOF Then

                For i = 0 To dict.Item("product")(0)("supplierItems").Count - 1
                    'Nur hinterlegte Distris ansprechen
                    Do Until rs.EOF
                        If rs.Item("id") = dict.Item("product")(0)("supplierItems")(i)("supplierId").ToString Then
                            supplier = New CSupplier
                            product.supplieritem.Add(supplier)
                            supplier.id = dict.Item("product")(0)("supplierItems")(i)("supplierId").ToString

                            Try
                                supplier.supplierItemId = dict.Item("product")(0)("supplierItems")(i)("supplierSKU").ToString

                            Catch ex As Exception
                            End Try

                            Try
                                supplier.stockStatusText = dict.Item("product")(0)("supplierItems")(i)("stockStatusText").ToString
                            Catch ex As Exception
                            End Try

                            supplier.stock = dict.Item("product")(0)("supplierItems")(i)("stock").ToString
                            supplier.price = Replace(dict.Item("product")(0)("supplierItems")(i)("price").ToString, ".", ",")
                            supplier.priceSourceId = 0
                            supplier.calcprice = ConversionHelper.ToDecimal(dict.Item("product")(0)("supplierItems")(i)("priceCalc").ToString)
                        End If
            rs.MoveNext()
                    Loop
                    rs.MoveFirst()
                Next

            End If

            rs.Close()

            reader.Close()
            DataStream.Close()
            Response.Close()

            Return product

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
            Return Nothing
        End Try

    End Function

End Module
