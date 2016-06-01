Imports System.Console

Module Program

    Sub Main()

        Dim args() As String = Environment.GetCommandLineArgs
        Dim puid As String = ""

        Dim connector As New Schwindt.OL2ITscope.Connector
        Dim Article As String = ""

        WriteLine("Schwindt Office Line To ITscope IMPORT-TOOL")
        WriteLine(Now)
        WriteLine()

        Dim test(0) As String
        'test.SetValue("schwindt://2275956000/", 0)

        Try

            WriteLine("ARTIKELIMPORT Test")
            WriteLine("--------------------------------------------------------------------------------")

            'WriteLine(args(1))
            'puid = test(0).Substring(11, Len(test(0)) - 12)
            puid = args(1).Substring(11, Len(args(1)) - 12)

            ForegroundColor = ConsoleColor.Yellow
            WriteLine("ITscope-ID: " & puid)
            ResetColor()
            WriteLine()

            If Logon.gbAnmeldung() Then

                connector.InitMandantNET(Logon.goMandant)

                Dim erg As Schwindt.OL2ITscope.CProduct

                erg = connector.CreateArticlePuid(puid, True)

                If erg.errorMessage > "" Then
                    ForegroundColor = ConsoleColor.Red
                    WriteLine("FEHLER: " & erg.errorMessage)
                    ResetColor()
                End If

                WriteLine()
                If Sagede.Core.Tools.ConversionHelper.ToString(erg.Article) <> "" Then
                    ForegroundColor = ConsoleColor.Green
                    WriteLine("Artikel wurde erfolreich angelegt: " & erg.Article)
                    ResetColor()
                    WriteLine()
                End If
                WriteLine("--------------------------------------------------------------------------------")
                WriteLine("VORGANG ABGESCHLOSSEN!")

                gbAbmeldung()

                Threading.Thread.Sleep(1000)
                System.Environment.Exit(0)
            Else
                WriteLine("--------------------------------------------------------------------------------")
                WriteLine("VORGANG ABGESCHLOSSEN! - Mit Enter beenden!")
            End If

        Catch ex As Exception
            'WriteLine("- Keine Argumente angegeben -")
            ForegroundColor = ConsoleColor.Red
            WriteLine(ex.Message)
            WriteLine()
            WriteLine("--------------------------------------------------------------------------------")
            WriteLine("VORGANG ABGESCHLOSSEN! - Mit Enter beenden!")
            ResetColor()

            gbAbmeldung()

        End Try

        ReadLine()

    End Sub

End Module
