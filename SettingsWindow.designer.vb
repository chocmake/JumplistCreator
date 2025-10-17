Imports System.Windows.Forms

Partial Public Class SettingsWindow
    Private components As System.ComponentModel.IContainer
    Private DefaultLaunchActionMenu As DefaultLaunchActionDropdown
    Private maxSettingsWidth As Integer = 200

    Friend WithEvents LanguageDropdownLabel As Label
    Friend WithEvents LanguageDropdownMenu As LanguageDropdown
    Friend WithEvents ThemeDropdownLabel As Label
    Friend WithEvents ThemeDropdownMenu As ThemeDropdown
    Friend WithEvents JumplistMaxItemsLabel As Label
    Friend WithEvents JumplistMaxItemsValBox As NumericUpDown
    Friend WithEvents JumplistMaxItemsTimer As Timer
    Friend WithEvents DefaultLaunchActionLabel As Label
    Friend WithEvents MetaItemsDropdownLabel As Label
    Friend WithEvents MetaItemsDropdownMenu As MetaItemsDropdown
    Friend WithEvents MetaItemsUpdateCheckbox As CheckBox
    Friend WithEvents OkButton As Button

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.SuspendLayout()

        ' SettingsWindow
        Me.Name = "SettingsWindow"
        Me.Text = Lang.GetString("Settings")
        Me.Icon = FileUtilities.LoadEmbeddedIcon("Resources.Icons.Main.ico")
        Me.ClientSize = New System.Drawing.Size(400, 314)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.AcceptButton = Me.OkButton
        Me.Font = Globals.DefaultFont

        Me.LanguageDropdownLabel = New Label() With {
            .Name = "LanguageDropdownLabel",
            .TabIndex = 0,
            .Text = Lang.GetString("LangLabel")
        }
        Me.Controls.Add(Me.LanguageDropdownLabel)
        FormFunc.SetItemSizeCoords(Me, LanguageDropdownLabel, width:=maxSettingsWidth)

        InitializeLanguageDropdown() ' initialized here so its tab index gets parsed for positioning in correct order

        Me.ThemeDropdownLabel = New Label() With {
            .Name = "ThemeDropdownLabel",
            .TabIndex = 2,
            .Text = Lang.GetString("ThemeLabel")
        }
        Me.Controls.Add(Me.ThemeDropdownLabel)
        FormFunc.SetItemSizeCoords(Me, ThemeDropdownLabel, width:=maxSettingsWidth)

        InitializeThemeDropdown()

        Me.JumplistMaxItemsLabel = New Label() With {
            .Name = "JumplistMaxItemsLabel",
            .TabIndex = 4,
            .Text = Lang.GetString("MaxItemsLabel")
        }
        Me.Controls.Add(Me.JumplistMaxItemsLabel)
        FormFunc.SetItemSizeCoords(Me, JumplistMaxItemsLabel, width:=maxSettingsWidth)

        Me.JumplistMaxItemsTimer = New Timer(Me.components) With {
            .Interval = 700 ' ms
        }

        Me.JumplistMaxItemsValBox = New NumericUpDown() With {
            .Name = "JumplistMaxItemsValBox",
            .TabIndex = 5,
            .DecimalPlaces = 0,
            .Minimum = 1D,
            .Maximum = 100D,
            .Increment = 1D
        }
        Me.Controls.Add(Me.JumplistMaxItemsValBox)
        FormFunc.SetItemSizeCoords(Me, JumplistMaxItemsValBox)
        FormFunc.AlignControlWithLabel(Me, JumplistMaxItemsValBox, JumplistMaxItemsLabel)

        Me.DefaultLaunchActionLabel = New Label() With {
            .Name = "DefaultLaunchActionLabel",
            .TabIndex = 6,
            .Text = Lang.GetString("DefaultLaunchActionLabel")
        }
        Me.Controls.Add(Me.DefaultLaunchActionLabel)
        FormFunc.SetItemSizeCoords(Me, DefaultLaunchActionLabel, width:=maxSettingsWidth)
        Me.DefaultLaunchActionLabel.Top -= Me.DefaultLaunchActionLabel.Height * Math.Max(0, (DpiFactor - 1)) \ (2 * DpiFactor) ' offset control position as DPI increases to compensate for corrected preceding NumericUpDown control offset at high DPIs (see `FormFunc.AlignControlWithLabel()`)

        InitializeDefaultLaunchActionDropdown()

        Me.MetaItemsDropdownLabel = New Label() With {
            .Name = "MetaItemsDropdownLabel",
            .TabIndex = 8,
            .Text = Lang.GetString("MetaListSettingsVisLabel")
        }
        Me.Controls.Add(Me.MetaItemsDropdownLabel)
        FormFunc.SetItemSizeCoords(Me, MetaItemsDropdownLabel, width:=maxSettingsWidth)

        InitializeMetaItemsDropdown()

        Me.MetaItemsUpdateCheckbox = New CheckBox() With {
            .Name = "MetaItemsUpdateCheckbox",
            .TabIndex = 10,
            .Text = Lang.GetString("MetaListSettingsUpdateItemLabel")
        }
        Me.Controls.Add(Me.MetaItemsUpdateCheckbox)
        FormFunc.SetItemSizeCoords(Me, MetaItemsUpdateCheckbox, width:=Me.MetaItemsDropdownMenu.Width) ' while the default height leaves some vertical whitespace it's necessary for if the localized string has to wrap to a new line, else the second line doesn't display
        Me.MetaItemsUpdateCheckbox.Left = Me.MetaItemsDropdownMenu.Left

        Me.OkButton = New Button() With {
            .Name = "OkButton",
            .TabIndex = 12,
            .Text = Lang.GetString("OK"),
            .UseVisualStyleBackColor = True,
            .DialogResult = DialogResult.OK
        }
        Me.Controls.Add(Me.OkButton)
        FormFunc.SetItemSizeCoords(Me, OkButton, setAutoWidth:=True, alignRight:=True)

        ' SettingsWindow
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub InitializeLanguageDropdown()
        Me.LanguageDropdownMenu = New LanguageDropdown() With {
            .Name = "LanguageDropdownMenu",
            .TabIndex = 1
        }
        Me.Controls.Add(Me.LanguageDropdownMenu)
        FormFunc.SetItemSizeCoords(Me, LanguageDropdownMenu)
        FormFunc.AlignControlWithLabel(Me, LanguageDropdownMenu, LanguageDropdownLabel)
    End Sub

    Private Sub InitializeThemeDropdown()
        Me.ThemeDropdownMenu = New ThemeDropdown() With {
            .Name = "ThemeDropdownMenu",
            .TabIndex = 3
        }
        Me.Controls.Add(Me.ThemeDropdownMenu)
        FormFunc.SetItemSizeCoords(Me, ThemeDropdownMenu)
        FormFunc.AlignControlWithLabel(Me, ThemeDropdownMenu, ThemeDropdownLabel)
    End Sub

    Private Sub InitializeDefaultLaunchActionDropdown()
        Me.DefaultLaunchActionMenu = New DefaultLaunchActionDropdown() With {
            .Name = "DefaultLaunchActionMenu",
            .TabIndex = 7
        }
        Me.Controls.Add(Me.DefaultLaunchActionMenu)
        FormFunc.SetItemSizeCoords(Me, DefaultLaunchActionMenu)
        FormFunc.AlignControlWithLabel(Me, DefaultLaunchActionMenu, DefaultLaunchActionLabel)
    End Sub

    Private Sub InitializeMetaItemsDropdown()
        Me.MetaItemsDropdownMenu = New MetaItemsDropdown() With {
            .Name = "MetaItemsDropdownMenu",
            .TabIndex = 9
        }
        Me.Controls.Add(Me.MetaItemsDropdownMenu)
        FormFunc.SetItemSizeCoords(Me, MetaItemsDropdownMenu)
        FormFunc.AlignControlWithLabel(Me, MetaItemsDropdownMenu, MetaItemsDropdownLabel)
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

Public Class InnerSizeHelper
    Private ReadOnly container As Control
    Private ReadOnly inner As Control

    Public Sub New(container As Control, inner As Control)
        If container Is Nothing Then Throw New ArgumentNullException(NameOf(container))
        If inner Is Nothing Then Throw New ArgumentNullException(NameOf(inner))
        Me.container = container
        Me.inner = inner
    End Sub

    ' Forward size changes to the inner ComboBox control (otherwise size changes on container don't get set as expected)
    Public Sub SyncSize()
        Dim paddingLeft As Integer = 0
        Dim paddingTop As Integer = Math.Max(0, (container.Height - inner.Height) \ 2)
        Dim innerWidth As Integer = Math.Max(0, container.Width - paddingLeft)
        inner.Location = New Point(paddingLeft, paddingTop)
        inner.Size = New Size(innerWidth, inner.Height)
    End Sub
End Class