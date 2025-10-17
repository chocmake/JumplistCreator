Imports System.Windows.Forms

Public Class ThemeDropdown
    Inherits UserControl

    Private Shared ReadOnly ModeLookup As New Dictionary(Of String, Integer) From {
        {"Auto", 0},
        {"Light", 1},
        {"Dark", 2}
    }

    Public Event ThemeChanged(ByVal sender As Object, ByVal newMode As String)

    Public Property SelectedMode As String
        Get
            Dim idx = If(Me.combo Is Nothing, -1, Me.combo.SelectedIndex)
            Return IndexToString(idx)
        End Get
        Set(value As String)
            Dim idx As Integer = StringToIndex(value)
            If Me.combo IsNot Nothing Then
                If idx < 0 OrElse idx >= Me.combo.Items.Count Then idx = 0
                Me.combo.SelectedIndex = idx
            End If
        End Set
    End Property

    Public Sub Populate()
        If Me.combo Is Nothing Then Return

        Me.combo.Items.Clear()
        Me.combo.Items.AddRange(New Object() {
            Lang.GetString("ThemeAuto"),
            Lang.GetString("ThemeLight"),
            Lang.GetString("ThemeDark")})

        Dim stored As String = GetCurrentMode()
        Me.SelectedMode = stored
    End Sub

    Private Function StringToIndex(key As String) As Integer
        If Not String.IsNullOrEmpty(key) AndAlso ModeLookup.ContainsKey(key) Then
            Return ModeLookup(key)
        End If
        Return 0
    End Function

    Private Function IndexToString(idx As Integer) As String
        For Each kvp In ModeLookup
            If kvp.Value = idx Then Return kvp.Key
        Next
        Return "Auto"
    End Function

    Private Function GetCurrentMode() As String
        Dim iniVal = IniSettings.ReadValue("Theme")
        If String.IsNullOrEmpty(iniVal) Then Return "Auto"

        If ModeLookup.ContainsKey(iniVal) Then
            Return iniVal
        End If

        Return "Auto"
    End Function

    Private Sub SetCurrentMode(newMode As String)
        If String.IsNullOrEmpty(newMode) Then newMode = "Auto"
        If Not ModeLookup.ContainsKey(newMode) Then newMode = "Auto"

        If Me.combo IsNot Nothing Then
            Me.SelectedMode = newMode
        End If

        IniSettings.WriteValue("Theme", newMode)
    End Sub

    Private Sub ComboSelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim idx = Me.combo.SelectedIndex
        Dim mode As String = IndexToString(idx)

        If mode <> GetCurrentMode() Then
            SetCurrentMode(mode)
            RaiseEvent ThemeChanged(Me, mode)
        End If
    End Sub
End Class
