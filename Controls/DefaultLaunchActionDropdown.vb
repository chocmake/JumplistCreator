Imports System.IO
Imports System.Windows.Forms

Partial Class DefaultLaunchActionDropdown
    Inherits UserControl

    Public Sub New()
        InitializeComponent()
        InitializeControl()
    End Sub

    ' Avoiding using `Const` to allow function calls
    Private Shared ReadOnly itemDefault As String = Lang.GetString("DefaultLaunchActionItem")
    Private Shared ReadOnly itemSelect As String = Lang.GetString("BrowseForFileButton")
    Private CustomItemIndex As Integer = -1

    Private Sub InitializeControl()
        combo.Items.Clear()
        combo.Items.Add(itemDefault)
        combo.Items.Add(itemSelect)
        CustomItemIndex = -1

        Dim stored As String = IniSettings.ReadValue("DefaultLaunchAction")
        If String.IsNullOrWhiteSpace(stored) Then
            combo.SelectedIndex = 0 ' Default
        Else
            InsertCustomItem(stored)
        End If

        AddHandler combo.SelectedIndexChanged, AddressOf ComboSelectedIndexChanged
    End Sub

    Private Sub ComboSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idx As Integer = combo.SelectedIndex
        If idx = -1 Then
            Return
        End If

        Dim selectedText As String = combo.Items(idx).ToString()

        If selectedText = itemSelect Then
            HandleSelectFile()
            Return
        End If

        If selectedText = itemDefault Then
            ' Remove custom item if present
            If CustomItemIndex <> -1 Then
                ' Remove by index (subsequently Select index becomes 1 again)
                combo.Items.RemoveAt(CustomItemIndex)
                CustomItemIndex = -1
            End If
            IniSettings.WriteValue("DefaultLaunchAction") ' write empty value
            combo.SelectedIndex = 0
            Return
        End If
    End Sub

    Private Sub InsertCustomItem(inputPath As String)
        ' Insert custom path item between Default and Select items
        CustomItemIndex = 1

        Dim parsed As String
        Dim parts As ParsedJumplistEntry = ParseValueCommand(inputPath)
        Dim targetPath As String = parts.TargetPath
        Dim arguments As String = parts.Args
        Try
            parsed = Path.GetFileName(targetPath)
        Catch ex As Exception
            ' Fall back to truncation of characters, based on how many glyphs fit in max pixels, if Path.GetFileName() throws an exception due to illegal characters
            Dim font As Font = Control.DefaultFont
            Dim maxWidthPx As Single = 120
            parsed = "..." & StringParser.TruncateToWidth(targetPath, font, maxWidthPx, True)
        End Try

        combo.Items.Insert(CustomItemIndex, parsed)
        combo.SelectedIndex = CustomItemIndex
    End Sub

    Private Sub HandleSelectFile()
        ofd.FileName = String.Empty
        ofd.Filter = "All files (*.*)|*.*"
        ofd.Title = Lang.GetString("BrowseForFileWindow")

        If ofd.ShowDialog() = DialogResult.OK Then
            Dim path As String = ofd.FileName
            ' If custom exists then replace it, otherwise insert
            If CustomItemIndex <> -1 Then
                combo.Items(CustomItemIndex) = path
                combo.SelectedIndex = CustomItemIndex
            Else
                InsertCustomItem(path)
            End If

            IniSettings.WriteValue("DefaultLaunchAction",path)
        Else
            ' If canceling browse file dialog then revert selection to previous item
            If CustomItemIndex <> -1 Then
                combo.SelectedIndex = CustomItemIndex
            Else
                combo.SelectedIndex = 0
            End If
        End If
    End Sub
End Class
