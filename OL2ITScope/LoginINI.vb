Imports Sagede.Core.Tools
Imports Microsoft.Win32

Module LoginINI

    Private Datei As String = "Login.ini"
    Private Dateipfad As String

    Public Sub ICreateLoginINI(ByVal Benutzer As String, ByVal Kennwort As String)

        Dim ini As System.IO.StreamWriter
        Dim Dateipfad As String
        Dim Datenbank() As String

        Dim RegKey As String
        Dim RKRoot As RegistryKey
        Dim RKSub As RegistryKey

        'RegKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\28319529AA0DE2E44996E66E971086E4\InstallProperties"
        RegKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\097980A75F638C741B7838A3E0821CF1\InstallProperties"

        RKRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
        RKSub = RKRoot.OpenSubKey(RegKey)

        Dateipfad = RKSub.GetValue("InstallLocation").ToString & "schwindt\"

        Datenbank = Split(OL2ITScope.Connector._mandant.DatasetName, ";")
        'Dateipfad = Mid(OL2ITScope.Connector._mandant.ApplicationDevice.SourceConfiguration.ConnectionParameter, 1, Len(OL2ITScope.Connector._mandant.ApplicationDevice.SourceConfiguration.ConnectionParameter) - 6) & "schwindt\"

        Try
            ini = My.Computer.FileSystem.OpenTextFileWriter(Dateipfad & Datei, False)

            ini.WriteLine(ConversionHelper.ToString(Datenbank(0)))
            ini.WriteLine(OL2ITScope.Connector._mandant.Id)
            ini.WriteLine(Benutzer)
            ini.WriteLine(IEncodeString(Kennwort, "schwindt"))

            MsgBox("Die Login-Datei wurde erfolgreich erstellt!", MsgBoxStyle.Information, "Login-Datei")
        Catch ex As Exception
            MsgBox("Die Anmeldedaten konnten nicht in einer Datei gespeichert werden!", MsgBoxStyle.Critical, "LoginINI")
        End Try

        ini.Close()

    End Sub

    Public Function IReadLogindaten() As CLogin

        Dim ini As System.IO.StreamReader
        Dim oLogindaten As New CLogin

        Dim RegKey As String
        Dim RKRoot As RegistryKey
        Dim RKSub As RegistryKey

        RegKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\097980A75F638C741B7838A3E0821CF1\InstallProperties"

        RKRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
        RKSub = RKRoot.OpenSubKey(RegKey)

        Dateipfad = RKSub.GetValue("InstallLocation").ToString & "schwindt\"

        Try
            ini = My.Computer.FileSystem.OpenTextFileReader(Dateipfad & Datei)

            Do Until ini.EndOfStream
                oLogindaten.Datenbank = ini.ReadLine
                oLogindaten.Mandant = ini.ReadLine
                oLogindaten.Benutzer = ini.ReadLine
                oLogindaten.Kennwort = IDecodeString(ini.ReadLine, "schwindt")
            Loop

            ini.Close()

            Return oLogindaten
        Catch ex As Exception

        End Try

    End Function

End Module
