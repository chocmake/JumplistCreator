Imports System.ComponentModel
Imports System.Windows.Forms

Partial Public Class MetaItemsDropdown
    Inherits UserControl

    Public Event MetaItemsChanged(ByVal sender As Object, ByVal newMode As Integer)
    Private curSelection As Integer = 1

    Private ReadOnly ModeLookup As New Dictionary(Of Integer, String) From {
        {0, "MetaListSettingsVisHidden"},
        {1, "MetaListSettingsVisConditional"},
        {2, "MetaListSettingsVisAlways"}
    }

    Private Sub Populate()
        Try
            Dim current As Integer = GetCurrentMode()
            If current < 0 OrElse current > 2 Then
                current = 1
            End If
            curSelection = current
        Catch ex As Exception
            curSelection = 1
        End Try

        PopulateItemsFromCurrentMode()
        SetSelectedItemForMode(curSelection)
    End Sub

    ' If current mode is 0 (always hidden, configurable only via manual INI editing) then show all items
    ' Otherwise only display 1 and 2. This avoids user footgun via GUI in the scenario where the config items are hidden and they've also set a default launch action (preventing the program's config window from opening). In that scenario they'd otherwise have to disable that INI setting or call the config via CLI to open the config window.
    Private Sub PopulateItemsFromCurrentMode()
        Me.combo.Items.Clear()

        Dim currentValue As Integer
        Try
            currentValue = GetCurrentMode()
            If currentValue < 0 OrElse currentValue > 2 Then
                currentValue = curSelection
            End If
        Catch ex As Exception
            currentValue = curSelection
        End Try

        If currentValue = 0 Then
            AddModeItem(0)
            AddModeItem(1)
            AddModeItem(2)
        Else
            AddModeItem(1)
            AddModeItem(2)
        End If
    End Sub

    Private Sub AddModeItem(m As Integer)
        Dim display As String = Nothing
        Dim key As String = Nothing

        If ModeLookup.TryGetValue(m, key) Then
            Try
                display = LookupDisplayString(key)
            Catch ex As Exception
                display = key
            End Try
        End If

        Dim wrapper As New ComboBoxItem With {
            .Text = display,
            .Value = m
        }
        Me.combo.Items.Add(wrapper)
    End Sub

    Private Sub SetSelectedItemForMode(mode As Integer)
        For i As Integer = 0 To Me.combo.Items.Count - 1
            Dim it = TryCast(Me.combo.Items(i), ComboBoxItem)
            If it IsNot Nothing AndAlso it.Value = mode Then
                Me.combo.SelectedIndex = i
                Return
            End If
        Next

        For i As Integer = 0 To Me.combo.Items.Count - 1
            Dim it = TryCast(Me.combo.Items(i), ComboBoxItem)
            If it IsNot Nothing AndAlso it.Value = 1 Then
                Me.combo.SelectedIndex = i
                Return
            End If
        Next

        If Me.combo.Items.Count > 0 Then
            Me.combo.SelectedIndex = 0
        Else
            Me.combo.SelectedIndex = -1
        End If
    End Sub

    Private Sub ComboSelectedIndexChanged(sender As Object, e As EventArgs)
        Dim newMode As Integer = curSelection
        Dim it = TryCast(Me.combo.SelectedItem, ComboBoxItem)
        If it IsNot Nothing Then
            newMode = it.Value
        End If

        If newMode <> curSelection Then
            SetCurrentMode(newMode)
            RaiseEvent MetaItemsChanged(Me, newMode)
        End If
    End Sub

    Public Sub RefetchFromSource()
        Try
            curSelection = GetCurrentMode()
        Catch ex As Exception
        End Try

        PopulateItemsFromCurrentMode()
        SetSelectedItemForMode(curSelection)
    End Sub

    Private Class ComboBoxItem
        Public Property Text As String
        Public Property Value As Integer
        Public Overrides Function ToString() As String
            Return If(String.IsNullOrEmpty(Text), Value.ToString(), Text)
        End Function
    End Class

    Public Function GetCurrentMode() As Integer
        Dim iniVal = IniSettings.ReadValue("MetaItemsVisibility")
        If iniVal <> curSelection Then
            SetCurrentMode(iniVal)
        End If
        Return Globals.MetaSettings.Visibility
    End Function

    Private Sub SetCurrentMode(newMode As Integer)
        curSelection = newMode
        IniSettings.WriteValue("MetaItemsVisibility", curSelection)
        Globals.MetaSettings.Visibility = newMode
    End Sub

    Private Function LookupDisplayString(key As String) As String
        Return Lang.GetString(key)
    End Function
End Class
