Imports System.Linq
Imports System.Windows.Forms

Partial Public Class LanguageDropdown
    Inherits UserControl

    Public Event LangChanged(ByVal sender As Object, ByVal newVal As String)
    Public Property ExternalHandler As Action(Of String)
    Private curSelection As String

    Private Sub LanguageDropdownLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        Populate()
    End Sub

    Private Sub Populate()
        Dim items = Lang.GetLangs()
        combo.Items.Clear()
        If items IsNot Nothing Then
            For Each li In items
                combo.Items.Add(li)
            Next
        End If

        combo.DisplayMember = "Name"
        combo.ValueMember = "Lang"

        curSelection = Lang.GetCurrentLang()
        SelectCurrent()
    End Sub

    Private Sub SelectCurrent()
        If combo.Items.Count = 0 Then Return
        For i As Integer = 0 To combo.Items.Count - 1
            Dim li = TryCast(combo.Items(i), LangInfo)
            If li IsNot Nothing AndAlso String.Equals(li.Lang, curSelection, StringComparison.OrdinalIgnoreCase) Then
                combo.SelectedIndex = i
                Return
            End If
        Next
        ' Fallback (select first if no match)
        If combo.Items.Count > 0 Then combo.SelectedIndex = 0
    End Sub

    Private Sub ComboSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim li = TryCast(combo.SelectedItem, LangInfo)
        If li Is Nothing Then Return

        If curSelection <> li.Lang Then
            curSelection = li.Lang
            RaiseEvent LangChanged(Me, li.Lang)
            If ExternalHandler IsNot Nothing Then ExternalHandler.Invoke(li.Lang)
            WriteSelection(li.Lang)
        End If
    End Sub

    Private Sub WriteSelection(token As String)
        IniSettings.WriteValue("Language", token)
    End Sub

    Public Sub RefetchFromSource()
        Dim iniVal = IniSettings.ReadValue("Language")
        If iniVal <> curSelection AndAlso Not String.IsNullOrEmpty(iniVal) Then
            curSelection = iniVal
            SelectCurrent()
            RaiseEvent LangChanged(Me, curSelection)
            If ExternalHandler IsNot Nothing Then ExternalHandler.Invoke(curSelection)
        End If
    End Sub
End Class
