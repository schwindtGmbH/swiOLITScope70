<Microsoft.VisualBasic.ComClass()> Public Class Versionsinfo

    <System.Runtime.InteropServices.ComVisible(True)> _
    Public Function getVersioninfo(ByVal Mode As Int32, ByVal Anwendung As String, ByVal Version As String, ByVal VersionBis As String) As String()

        Return IgetVersionsinfo(Mode, Anwendung, Version, VersionBis)

    End Function

End Class
