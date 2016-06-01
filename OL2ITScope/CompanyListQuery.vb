Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Sagede.OfficeLine.Data
Imports Sagede.OfficeLine.Engine
Imports Sagede.Core.Tools

Module CompanyListQuery

    Public Function IReadCompanyList() As Boolean

        'Dim Request As HttpWebRequest = CType(WebRequest.Create("https://apitest.itscope.com/1.0/company/distributor/companylist.json"), HttpWebRequest)
        Dim Request As HttpWebRequest = CType(WebRequest.Create("https://api.itscope.com/2.0/company/distributor/company.json"), HttpWebRequest)
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
                Qry = "DELETE From swiITscopeDistributoren WHERE Mandant =" & Connector._mandant.Id & ""
                Connector._mandant.MainDevice.GenericConnection.ExecuteNonQuery(Qry)
            End If

            Qry = "SELECT * From swiITscopeDistributoren WHERE 1=2"
            rs = New GenericRecordset(Connector._mandant.MainDevice)
            rs.Open(Qry, False)

            For i = 0 To dict.Item("company").count - 1
                Try
                    rs.AddNew()

                    rs!id = (dict.Item("company")(i)("supplier")("id").ToString)
                    rs!name = (dict.Item("company")(i)("name").ToString)
                    rs!mandant = Connector._mandant.Id

                    rs.Update()
                    '30.10.2015, MG
                    'Catch ex As Exception
                    '    rs.Close()
                    '    
                    '    If Err.Number = 5 Then
                    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Fehler")
                    '        Return False
                    '    End If

                    '    rs = New GenericRecordset(Connector._mandant.MainDevice)
                    '    rs.Open(Qry, False)
                Catch ex As Sagede.OfficeLine.Data.RecordsetException
                    rs.Close()
                    If ex.ErrorNumber = "-2146233088" Then
                        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Fehler")
                        Return False
                    End If
                Catch ex2 As Exception
                    rs.Close()

                    rs = New GenericRecordset(Connector._mandant.MainDevice)
                    rs.Open(Qry, False)
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
