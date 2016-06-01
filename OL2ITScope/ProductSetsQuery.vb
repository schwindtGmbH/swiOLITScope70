Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Sagede.OfficeLine.Data
Imports Sagede.OfficeLine.Engine
Imports Sagede.Core.Tools

Module ProductSetsQuery

    Public Function IReadProductSets() As Boolean

        'Dim Request As HttpWebRequest = CType(WebRequest.Create("https://apitest.itscope.com/1.0/products/sets/setlist.json"), HttpWebRequest)
        Dim Request As HttpWebRequest = CType(WebRequest.Create("https://api.itscope.com/2.0/products/producttypes/producttype.json"), HttpWebRequest)
        Dim Login64Base As String = GetLogin64Base()
        Dim result As MsgBoxResult

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
            Dim Qry As String
            Dim rs As IGenericRecordset

            result = MsgBox("Wollen Sie die Einstellungen neu importieren?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Import")
            If result = MsgBoxResult.Yes Then
                Qry = "DELETE From swiITscopeArtikelgruppen WHERE Mandant = " & Connector._mandant.Id & ""
                Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(Qry)
            End If

            Qry = "SELECT * From swiITscopeArtikelgruppen WHERE 1=2"
            rs = New GenericRecordset(Connector._mandant.MainDevice)
            rs.Open(Qry, False)

            For i = 0 To dict.Item("productType").count - 1
                Try
                    Try
                        rs.AddNew()
                    Catch ex As Exception

                    End Try

                    rs!id = (dict.Item("productType")(i)("id").ToString)
                    rs!name = (dict.Item("productType")(i)("name").ToString)
                    rs!groupid = (dict.Item("productType")(i)("productTypeGroup")("id").ToString)
                    rs!groupname = (dict.Item("productType")(i)("productTypeGroup")("name").ToString)
                    rs!mandant = Connector._mandant.Id

                    rs.Update()
                Catch ex As Exception

                End Try

            Next

            rs.Close()

            reader.Close()
            DataStream.Close()
            Response.Close()

            Return True

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Fehler")
            Return False
        End Try

    End Function

End Module
