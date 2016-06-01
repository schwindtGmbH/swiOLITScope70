Public Class frmArtikelAnlage

    Private Sub cmdAnlage_Click(sender As System.Object, e As System.EventArgs) Handles cmdAnlage.Click

        If Me.txtArtikel.Text <> Me.txtTemp.Text Then
            Me.Close()
        Else
            MsgBox("Die Artikelnummer muss von der Hersteller-Artikelnr. abweichen, da der Artikel bereits vorhanden ist!", MsgBoxStyle.Exclamation, "Fehler Artikelnr.")
        End If

    End Sub

    Private Sub txtArtikel_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtArtikel.TextChanged

        If CheckArtikelnummer(Me.txtArtikel.Text) = False Then
            Me.cmdAnlage.Enabled = True
        Else
            Me.cmdAnlage.Enabled = False
        End If

    End Sub

    Private Sub frmArtikelAnlage_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.cmdAnlage.Enabled = False
    End Sub

    Private Sub cmdLink_Click(sender As System.Object, e As System.EventArgs) Handles cmdLink.Click

        Dim result As MsgBoxResult
        result = MsgBox("Wollen Sie den vorhandenen Artikel mit ITscope verknüpfen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Verknüpfen")
        If result = MsgBoxResult.Yes Then
            SetPuid4Article(Me.txtPuid.Text, Me.txtArtikel.Text)
            Me.txtTemp.Text = "Link"
            Me.Close()
        End If

    End Sub

End Class