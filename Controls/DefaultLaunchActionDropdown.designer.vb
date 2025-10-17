Imports System.Windows.Forms

Partial Class DefaultLaunchActionDropdown
    Inherits UserControl

    Private components As System.ComponentModel.IContainer
    Friend WithEvents combo As ComboBox
    Friend WithEvents ofd As OpenFileDialog
    Friend WithEvents helper As InnerSizeHelper

    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(disposing As Boolean)
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
        Me.components = New System.ComponentModel.Container()
        Me.combo = New ComboBox()
        Me.ofd = New OpenFileDialog()
        Me.helper = New InnerSizeHelper(Me, combo)
        Me.SuspendLayout()

        Me.combo.Anchor = CType(((AnchorStyles.Top Or AnchorStyles.Left) _
            Or AnchorStyles.Right), AnchorStyles)
        Me.combo.DropDownStyle = ComboBoxStyle.DropDownList
        Me.combo.FormattingEnabled = True
        Me.combo.Location = New System.Drawing.Point(0, 0)
        Me.combo.Name = "combo"
        Me.combo.Size = New System.Drawing.Size(220, Globals.ComboBoxDefaultHeight)
        Me.combo.TabIndex = 0

        Me.AutoScaleMode = AutoScaleMode.None
        Me.Controls.Add(Me.combo)
        Me.Name = "DefaultLaunchActionDropdown"
        Me.Size = New System.Drawing.Size(220, Globals.ComboBoxDefaultHeight)
        Me.ResumeLayout(False)
    End Sub

    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        Me.helper.SyncSize()
    End Sub
End Class