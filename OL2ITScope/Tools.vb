Module Tools

    Public Sub AusgabeInTextFile(ByVal Meldung As String)

        Dim txt As System.IO.StreamWriter

        If Not Connector._mandant.PropertyManager.GetValue(74914) = "0" Then

            Try
                txt = My.Computer.FileSystem.OpenTextFileWriter(Connector._mandant.PropertyManager.GetValue(74914) & "ITscope_Meldungen.txt", True)

                txt.WriteLine(Now)

                txt.WriteLine(Meldung)
                txt.WriteLine("-----------------------------------------------------------------------")

                txt.Close()
            Catch ex As Exception
                MsgBox("Protokoll konnte nicht geöffnet werden!", MsgBoxStyle.Critical, "Öffnen Protokoll")
            End Try

        End If

    End Sub

    Public Function GetVersion() As String

        Dim oAssembly = System.Reflection.Assembly.GetExecutingAssembly().GetName
        Dim version = oAssembly.Version.ToString

        Return version

    End Function

End Module
