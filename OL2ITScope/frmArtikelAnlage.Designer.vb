<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmArtikelAnlage
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmArtikelAnlage))
        Me.cmdAnlage = New System.Windows.Forms.Button()
        Me.lblArtikel = New System.Windows.Forms.Label()
        Me.txtArtikel = New System.Windows.Forms.TextBox()
        Me.lblHerstellerArtikelnr = New System.Windows.Forms.Label()
        Me.txtTemp = New System.Windows.Forms.TextBox()
        Me.cmdLink = New System.Windows.Forms.Button()
        Me.txtPuid = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'cmdAnlage
        '
        Me.cmdAnlage.Location = New System.Drawing.Point(178, 60)
        Me.cmdAnlage.Name = "cmdAnlage"
        Me.cmdAnlage.Size = New System.Drawing.Size(75, 23)
        Me.cmdAnlage.TabIndex = 0
        Me.cmdAnlage.Text = "&Anlage"
        Me.cmdAnlage.UseVisualStyleBackColor = True
        '
        'lblArtikel
        '
        Me.lblArtikel.AutoSize = True
        Me.lblArtikel.Location = New System.Drawing.Point(12, 36)
        Me.lblArtikel.Name = "lblArtikel"
        Me.lblArtikel.Size = New System.Drawing.Size(73, 13)
        Me.lblArtikel.TabIndex = 1
        Me.lblArtikel.Text = "Artikelnummer"
        '
        'txtArtikel
        '
        Me.txtArtikel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtArtikel.Location = New System.Drawing.Point(97, 34)
        Me.txtArtikel.Name = "txtArtikel"
        Me.txtArtikel.Size = New System.Drawing.Size(156, 20)
        Me.txtArtikel.TabIndex = 2
        '
        'lblHerstellerArtikelnr
        '
        Me.lblHerstellerArtikelnr.AutoSize = True
        Me.lblHerstellerArtikelnr.Location = New System.Drawing.Point(12, 9)
        Me.lblHerstellerArtikelnr.Name = "lblHerstellerArtikelnr"
        Me.lblHerstellerArtikelnr.Size = New System.Drawing.Size(98, 13)
        Me.lblHerstellerArtikelnr.TabIndex = 3
        Me.lblHerstellerArtikelnr.Text = "Hersteller-Artikelnr.:"
        '
        'txtTemp
        '
        Me.txtTemp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTemp.Location = New System.Drawing.Point(15, 63)
        Me.txtTemp.Name = "txtTemp"
        Me.txtTemp.Size = New System.Drawing.Size(45, 20)
        Me.txtTemp.TabIndex = 4
        Me.txtTemp.Visible = False
        '
        'cmdLink
        '
        Me.cmdLink.Location = New System.Drawing.Point(97, 60)
        Me.cmdLink.Name = "cmdLink"
        Me.cmdLink.Size = New System.Drawing.Size(75, 23)
        Me.cmdLink.TabIndex = 5
        Me.cmdLink.Text = "&Verknüpfen"
        Me.cmdLink.UseVisualStyleBackColor = True
        '
        'txtPuid
        '
        Me.txtPuid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPuid.Location = New System.Drawing.Point(40, 63)
        Me.txtPuid.Name = "txtPuid"
        Me.txtPuid.Size = New System.Drawing.Size(45, 20)
        Me.txtPuid.TabIndex = 6
        Me.txtPuid.Visible = False
        '
        'frmArtikelAnlage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(267, 93)
        Me.Controls.Add(Me.txtPuid)
        Me.Controls.Add(Me.cmdLink)
        Me.Controls.Add(Me.txtTemp)
        Me.Controls.Add(Me.lblHerstellerArtikelnr)
        Me.Controls.Add(Me.txtArtikel)
        Me.Controls.Add(Me.lblArtikel)
        Me.Controls.Add(Me.cmdAnlage)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmArtikelAnlage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ITscope Artikel-Anlage"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdAnlage As System.Windows.Forms.Button
    Friend WithEvents lblArtikel As System.Windows.Forms.Label
    Friend WithEvents txtArtikel As System.Windows.Forms.TextBox
    Friend WithEvents lblHerstellerArtikelnr As System.Windows.Forms.Label
    Friend WithEvents txtTemp As System.Windows.Forms.TextBox
    Friend WithEvents cmdLink As System.Windows.Forms.Button
    Friend WithEvents txtPuid As System.Windows.Forms.TextBox
End Class
