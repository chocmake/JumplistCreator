Imports System.Windows.Forms

Partial Public Class MetaItemsDropdown
    Inherits UserControl

    Private components As System.ComponentModel.IContainer
    Friend WithEvents combo As ComboBox
    Friend WithEvents helper As InnerSizeHelper

    Public Sub New()
        InitializeComponent()
        Populate()
        AddHandler Me.combo.SelectedIndexChanged, AddressOf ComboSelectedIndexChanged
    End Sub

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.combo = New ComboBox()
        Me.helper = New InnerSizeHelper(Me, combo)
        Me.SuspendLayout()

        Me.combo.DropDownStyle = ComboBoxStyle.DropDownList
        Me.combo.FormattingEnabled = True
        Me.combo.Location = New System.Drawing.Point(0, 0)
        Me.combo.Name = "combo"
        Me.combo.Size = New System.Drawing.Size(200, Globals.ComboBoxDefaultHeight)
        Me.combo.TabIndex = 0

        Me.Controls.Add(Me.combo)

        Me.Name = "MetaItemsDropdown"
        Me.AutoScaleMode = AutoScaleMode.None
        Me.Size = New System.Drawing.Size(200, Globals.ComboBoxDefaultHeight)
        Me.ResumeLayout(False)
    End Sub

    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        Me.helper.SyncSize()
    End Sub
End Class
