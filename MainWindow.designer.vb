'
' Updated by chocmake (https://github.com/chocmake) from original `Form1.designer.vb` beginning 2025-07-20
'

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainWindow
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    Friend WithEvents EditJumplistButton As Button
    Friend WithEvents UpdateJumplistButton As Button
    Friend WithEvents SettingsButton As ImageMaskButton
    Friend WithEvents AboutButton As ImageMaskButton
    Friend WithEvents UpdateProgressBar As RoundedProgressBar
    Friend WithEvents UpdateProgressTimer As Timer

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

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Globals.DpiFactor = DpiParser.GetFactor(Me)
        Me.components = New System.ComponentModel.Container()
        Me.SuspendLayout()
        
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainWindow))
        Dim windowWidth As Integer = 200
        Dim bottomButtonDim As Integer = 16

        ' MainWindow
        Me.Name = "MainWindow"
        Me.Text = Globals.ProgramName
        Me.ClientSize = New System.Drawing.Size(windowWidth, 136)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Icon = FileUtilities.LoadEmbeddedIcon("Resources.Icons.Main.ico")
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MinimizeBox = False ' Setting this together with `MaximizeBox` to false hides both buttons, allowing the title to not be truncated at narrow window widths
        Me.MaximizeBox = False
        Me.Font = Globals.DefaultFont

        Me.EditJumplistButton = New Button() With {
            .Name = "EditJumplistButton",
            .TabIndex = 0,
            .Text = Lang.GetString("EditJumplist"),
            .UseVisualStyleBackColor = True
        }
        Me.Controls.Add(Me.EditJumplistButton)
        FormFunc.SetItemSizeCoords(Me, EditJumplistButton)
        
        Me.UpdateJumplistButton = New Button() With {
            .Name = "UpdateJumplistButton",
            .TabIndex = 1,
            .Text = Lang.GetString("UpdateJumplist"),
            .UseVisualStyleBackColor = True
        }
        Me.Controls.Add(Me.UpdateJumplistButton)
        FormFunc.SetItemSizeCoords(Me, UpdateJumplistButton)
        
        Me.SettingsButton = New ImageMaskButton() With {
            .Name = "SettingsButton",
            .TabIndex = 2,
            .Size = New Size(bottomButtonDim, bottomButtonDim),
            .Text = "",
            .MaskImage = If(DpiFactor = 1,
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Icon-Settings-Mask.png"),
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Icon-Settings-Mask-2x.png")),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        Me.Controls.Add(Me.SettingsButton)
        FormFunc.SetItemSizeCoords(Me, SettingsButton, setSize:=False)
        Me.SettingsButton.Top += (MARGIN_GENERIC - 5)
        Me.SettingsButton.Left += 2

        Me.AboutButton = New ImageMaskButton() With {
            .Name = "AboutButton",
            .TabIndex = 3,
            .Size = New Size(bottomButtonDim, bottomButtonDim),
            .Text = "",
            .MaskImage = If(DpiFactor = 1,
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Icon-About-Mask.png"),
                           FileUtilities.LoadEmbeddedImage("Resources.Images.Icon-About-Mask-2x.png")),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        Me.Controls.Add(Me.AboutButton)
        FormFunc.SetItemSizeCoords(Me, AboutButton, setSize:=False)
        Me.AboutButton.Top = Me.SettingsButton.Top
        Me.AboutButton.Left = Me.SettingsButton.Left + bottomButtonDim + MARGIN_GENERIC

        Me.UpdateProgressBar = New RoundedProgressBar() With {
            .Name = "UpdateProgressBar",
            .TabIndex = 4,
            .Value = 0,
            .Minimum = 0,
            .Maximum = 100,
            .Radius = 2 * DpiFactor,
            .BorderWidth = 1 * DpiFactor,
            .BarBackColor = Color.Transparent
        } ' was anchored to bottom-right but shifted unexpectedly at 175-200% DPI
        Me.Controls.Add(Me.UpdateProgressBar)
        ' Setting height statically instead of being based on SettingsButton height to avoid inconsistencies at higher DPIs
        FormFunc.SetItemSizeCoords(Me, UpdateProgressBar, width:=44, height:=bottomButtonDim + Me.UpdateProgressBar.BorderWidth * 2)
        Me.UpdateProgressBar.Top = Me.SettingsButton.Top - 1
        Me.UpdateProgressBar.Left = Me.UpdateJumplistButton.Right - Me.UpdateProgressBar.Width - 2

        Me.UpdateProgressTimer = New Timer() With {
            .Interval = 20,
            .Enabled = False
        }

        ' Colors
        If ThemeLoader.DarkThemeIsActive() Then
            Me.SettingsButton.ColorNormal = Color.FromArgb(110, 110, 110)
            Me.SettingsButton.ColorHover = Color.FromArgb(150, 150, 150)
            Me.SettingsButton.ColorPressed = Color.FromArgb(230,230,230)
            Me.SettingsButton.ColorFocused = Me.SettingsButton.ColorHover
            Me.SettingsButton.ColorDisabled = Color.FromArgb(110, 110, 110)

            Me.UpdateProgressBar.BorderColor = Color.FromArgb(130, 110, 110, 110)
            Me.UpdateProgressBar.BarColor = Color.FromArgb(0, 107, 50)
        Else
            Me.SettingsButton.ColorNormal = Color.FromArgb(135, 135, 135)
            Me.SettingsButton.ColorHover = Color.FromArgb(90, 90, 90)
            Me.SettingsButton.ColorPressed = Color.FromArgb(30,30,30)
            Me.SettingsButton.ColorFocused = Me.SettingsButton.ColorHover
            Me.SettingsButton.ColorDisabled = Color.FromArgb(135, 135, 135)

            Me.UpdateProgressBar.BorderColor = Color.FromArgb(150, 135, 135, 135)
            Me.UpdateProgressBar.BarColor = Color.FromArgb(65, 191, 124)
        End If

        Me.AboutButton.ColorNormal = Me.SettingsButton.ColorNormal
        Me.AboutButton.ColorHover = Me.SettingsButton.ColorHover
        Me.AboutButton.ColorPressed = Me.SettingsButton.ColorPressed
        Me.AboutButton.ColorFocused = Me.SettingsButton.ColorHover
        Me.AboutButton.ColorDisabled = Me.SettingsButton.ColorDisabled

        ' MainWindow
        Me.ResumeLayout(False)
        Me.PerformLayout()
        Me.Width = (windowWidth + 16) * DpiFactor ' blunt kludge as workaround for incorrect width at 175-200% DPI but not at 100-150% (can't pinpoint cause)
    End Sub

    Private Sub StyleForUpdateArg()
        Me.UpdateProgressBar.Anchor = AnchorStyles.Bottom
        Me.Height = 83 * DpiFactor ' shorten to display only bottom controls
        EditJumplistButton.Visible = False
        EditJumplistButton.Enabled = False
        UpdateJumplistButton.Visible = False
        UpdateJumplistButton.Enabled = False
        SettingsButton.Visible = False
        SettingsButton.Enabled = False
        AboutButton.Visible = False
        AboutButton.Enabled = False
        UpdateProgressBar.Width = UpdateProgressBar.Right - SettingsButton.Left
        UpdateProgressBar.Left = SettingsButton.Left
        Me.Refresh() ' ensure a complete paint before calling UpdateJumplist()
    End Sub
End Class

Public Module FormFunc
    Private Const FONT_HEIGHT_REPL = 16 ' defined height statically for 96DPI (default) instead of based on `control.Font.Height` to avoid inconsistency of metrics at higher DPIs

    Private Sub AdjustForDarkTheme(control As Control)
        If TypeOf control Is Button AndAlso ThemeLoader.DarkThemeIsActive() Then
            ' Compensate for adjusted offsets/dimensions of DarkModeForms' rounded button styling. Tested alternatively applying this compensation via DarkModeForms itself but while the first buttons were correct subsequent buttons were off by -1px on Y.
            control.Top -= 2
            control.Left -= 2
            control.Width += 3
            control.Height += 3
        End If
    End Sub

    Public Sub SetItemSizeCoords(
        container As Control,
        control As Control, 
        Optional width As Integer? = Nothing,
        Optional height As Integer? = Nothing,
        Optional setSize As Boolean = True,
        Optional setLocation As Boolean = True,
        Optional setAutoWidth As Boolean = False,
        Optional alignRight As Boolean = False)

        If setSize Then
            Dim controlHeight As Integer = FONT_HEIGHT_REPL + If(TypeOf control Is NumericUpDown, 10, 20)
            Dim strSize As Size = Nothing

            If setAutoWidth Then
                control.Padding = New Padding((MARGIN_GENERIC * 4), 0, (MARGIN_GENERIC * 4), 0)
                strSize = control.GetPreferredSize(New Size(0, 0))
            End If

            control.Size = New Size(
                If(width.HasValue,
                   width.Value,
                   If(setAutoWidth,
                      strSize.Width,
                      container.ClientSize.Width - (2 * MARGIN_GENERIC))
                    ),
                If(height.HasValue,
                   height.Value,
                   controlHeight)
            )
        End If
        
        If setLocation Then
            ' Align first control (TabIndex 0) differently
            If control.TabIndex = 0 Then
                control.Location = New Point(MARGIN_GENERIC, MARGIN_GENERIC)
                AdjustForDarkTheme(control)
                Return
            End If

            ' Find bottom of last control
            Dim lastControl = container.Controls.Cast(Of Control)() _
                .Where(Function(c) c.Bottom > 0 AndAlso c.TabIndex < control.TabIndex) _
                .OrderByDescending(Function(c) c.Bottom) _
                .FirstOrDefault()

            Dim xPosition As Integer = If(alignRight,
                container.ClientSize.Width - control.Width - MARGIN_GENERIC,
                MARGIN_GENERIC)

            Dim yPosition As Integer = If(lastControl Is Nothing, 
                MARGIN_GENERIC, 
                lastControl.Bottom + (MARGIN_GENERIC \ 2))
            
            control.Location = New Point(xPosition, yPosition)
        End If
        AdjustForDarkTheme(control)
    End Sub

    Public Sub AlignControlWithLabel(
        container As Control,
        control As Control,
        label As Label)
        label.TextAlign = ContentAlignment.MiddleLeft
        
        Dim labelVerticalCenter As Integer = label.Top + (label.Height \ 2)
        
        Dim height As Integer = control.Height
        If height <= 0 Then height = FONT_HEIGHT_REPL + 8

        Dim extraFactor = If(TypeOf control Is NumericUpDown, DpiFactor, 1) ' at higher DPIs NumericUpDown specifically offsets oddly so this adjusts for it based on DPI factor
        Dim controlVerticalCenter As Integer = labelVerticalCenter - (height \ (2 * extraFactor))

        control.Location = New Point(label.Right + MARGIN_GENERIC, controlVerticalCenter)

        Dim newWidth As Integer = Math.Max(20, container.ClientSize.Width - label.Width - (3 * MARGIN_GENERIC))
        control.Size = New Size(newWidth, height)
    End Sub
End Module