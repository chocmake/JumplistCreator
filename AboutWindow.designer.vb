Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing

Partial Public Class AboutWindow
    Private components As System.ComponentModel.IContainer
    Friend WithEvents ProgramIcon As PictureBox
    Friend WithEvents LogoText As ImageMaskButton
    Friend WithEvents VersionLabel As Label
    Friend WithEvents WebsiteLinkLabel As LinkLabel
    Friend WithEvents CloseButton As Button

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.SuspendLayout()

        ' AboutWindow
        Me.Name = "AboutWindow"
        Me.Text = Lang.GetString("About")
        Me.Icon = FileUtilities.LoadEmbeddedIcon("Resources.Icons.Main.ico")
        Me.ClientSize = New System.Drawing.Size(417, 207)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.AcceptButton = Me.CloseButton
        Me.Font = Globals.DefaultFont

        Me.ProgramIcon = New PictureBox() With {
            .Name = "ProgramIcon",
            .Image = If(DpiFactor = 1,
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Program-Icon.png"),
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Program-Icon-2x.png")),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .Size = New Size(78, 78),
            .Top = 44,
            .Left = 87
        }
        Me.Controls.Add(ProgramIcon)

        Me.LogoText = New ImageMaskButton() With {
            .Name = "LogoText",
            .Size = New Size(148, 19),
            .Top = 55,
            .Left = 169,
            .Text = "",
            .MaskImage = If(DpiFactor = 1,
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Logo-Mask.png"),
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Logo-Mask-2x.png"))
        }
        Me.Controls.Add(Me.LogoText)

        Me.VersionLabel = New Label() With {
            .Name = "VersionLabel",
            .Height = 14 / DpiFactor,
            .Top = 77,
            .Left = 166,
            .Text = Lang.GetString("Version") & " " & Globals.VersionNumber
        } ' height divided by DPI factor to prevent it growing and its intrinsic background fill overlapping and obscuring the following LinkLabel at higher DPIs
        Me.Controls.Add(Me.VersionLabel)

        Me.WebsiteLinkLabel = New LinkLabel() With {
            .Name = "WebsiteLinkLabel",
            .TabIndex = 0,
            .Top = 95,
            .Left = 166,
            .Text = Lang.GetString("Website"),
            .LinkBehavior = LinkBehavior.HoverUnderline
        }
        Me.WebsiteLinkLabel.Links.Add(0, Me.WebsiteLinkLabel.Text.Length, Globals.WebsiteUrl)
        Me.Controls.Add(Me.WebsiteLinkLabel)

        Me.CloseButton = New Button() With {
            .Name = "CloseButton",
            .TabIndex = 1,
            .Text = Lang.GetString("Close"),
            .UseVisualStyleBackColor = True,
            .DialogResult = DialogResult.OK
        }
        Me.Controls.Add(Me.CloseButton)
        FormFunc.SetItemSizeCoords(Me, CloseButton, setAutoWidth:=True, alignRight:=True)
        Me.CloseButton.Top = Me.ClientSize.Height - Me.CloseButton.Height - MARGIN_GENERIC

        If ThemeLoader.DarkThemeIsActive() Then
            Me.LogoText.ColorNormal = Color.FromArgb(255, 255, 255)
            Me.WebsiteLinkLabel.LinkColor = Color.FromArgb(132, 173, 255)
        Else
            Me.LogoText.ColorNormal = Color.FromArgb(10, 10, 10)
            Me.WebsiteLinkLabel.LinkColor = Color.FromArgb(66, 149, 242)
        End If
        Me.LogoText.ColorHover = Me.LogoText.ColorNormal
        Me.LogoText.ColorPressed = Me.LogoText.ColorNormal
        Me.LogoText.ColorFocused = Me.LogoText.ColorNormal
        Me.LogoText.ColorDisabled = Me.LogoText.ColorNormal
         Me.WebsiteLinkLabel.ActiveLinkColor = Me.WebsiteLinkLabel.LinkColor
         Me.WebsiteLinkLabel.VisitedLinkColor = Me.WebsiteLinkLabel.LinkColor

        ' AboutWindow
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class