Imports Sagede.OfficeLine.Data
Imports Sagede.OfficeLine.Engine
Imports Sagede.OfficeLine.Shared
Imports Sagede.Core.Tools
Imports System.Console

Module Logon

    Private _session As Sagede.OfficeLine.Engine.Session
    Public goMandant As Sagede.OfficeLine.Engine.Mandant = Nothing

    Public Function gbAnmeldung() As Boolean

        Dim oLogin As New Schwindt.OL2ITscope.CLogin

        WriteLine("Office Line Anmeldung wird vorbereitet...!")
        WriteLine()

        oLogin = oLogin.ReadLogin

        If oLogin.Datenbank = "" Then
            WriteLine("Keine Logindaten vorhanden - bitte über Konfiguration INI-Datei erstellen!")
            WriteLine()
            Return False
        End If

        If goMandant Is Nothing Then

            Try
                _session = ApplicationEngine.CreateSession(oLogin.Datenbank, ApplicationToken.Abf, Nothing, New NamePasswordCredential(oLogin.Benutzer, oLogin.Kennwort))
                goMandant = _session.CreateMandant(oLogin.Mandant)

                WriteLine("Office Line Anmeldung erfolgreich!")
                WriteLine()

            Catch ex As Exception
                WriteLine("Office Line Anmeldung fehlgeschlagen!")
                WriteLine(ex.Message)
                WriteLine()

                Return False
                Exit Function

            End Try

        End If

        Return True

    End Function

    Public Function gbAbmeldung() As Boolean

        goMandant = Nothing
        gbAbmeldung = True

        WriteLine()
        WriteLine("ABMELDUNG Office Line erfolgreich!")
        WriteLine()

    End Function

End Module
