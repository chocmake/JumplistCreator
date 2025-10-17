Partial Public Class SettingsWindow
    Inherits System.Windows.Forms.Form

    Private pendingActions As Dictionary(Of String, Boolean) = New Dictionary(Of String, Boolean)()

    Public Sub New()
        InitializeComponent()

        ' Set defaults
        pendingActions("UpdateJumplist") = False
        pendingActions("RelaunchProgram") = False
        JumplistMaxItemsValBox.Text = Globals.JumpListItemsMaximum
        
        If Not ThemeLoader.ThemeLibExists() Then
            ThemeDropdownMenu.Enabled = False
        End If
        MetaItemsUpdateShouldDisable()

        ' Apply theme only after adjusting any initial enabled states
        ThemeLoader.TryApplyTheme(Me)
        
        AddHandler Me.Load, AddressOf SettingsWindowLoad
        AddHandler Me.FormClosed, AddressOf SettingsWindowClosed
        AddHandler Me.JumplistMaxItemsValBox.KeyPress, AddressOf Me.JumplistMaxItemsValBoxKeyPress
        AddHandler Me.JumplistMaxItemsValBox.ValueChanged, AddressOf Me.JumplistMaxItemsValBoxValueChanged
        AddHandler Me.JumplistMaxItemsTimer.Tick, AddressOf Me.JumplistMaxItemsTimerTick
        AddHandler Me.MetaItemsUpdateCheckbox.CheckedChanged, AddressOf MetaItemsUpdateCheckboxToggled
        AddHandler Me.OkButton.Click, AddressOf Me.OkButtonClick

        AddHandler LanguageDropdownMenu.LangChanged, AddressOf LanguageDropdownMenuSelChanged
        AddHandler ThemeDropdownMenu.ThemeChanged, AddressOf ThemeDropdownMenuSelChanged
        AddHandler MetaItemsDropdownMenu.MetaItemsChanged, AddressOf MetaItemsDropdownMenuModeChanged
        AddHandler DefaultLaunchActionMenu.DefaultLaunchActionChanged, AddressOf DefaultLaunchActionMenuChanged
    End Sub

    Private Sub OkButtonClick(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub SettingsWindowLoad(sender As Object, e As EventArgs)
        ' Reread settings from INI on load (for controls that don't do it already) in case they've been changed manually in-between window opens
        Me.MetaItemsDropdownMenu.RefetchFromSource()
        Me.LanguageDropdownMenu.RefetchFromSource()
        MetaItemsUpdateRefetch()
    End Sub

    Private Sub SettingsWindowClosed(sender As Object, e As EventArgs)
        RunIfQueued("UpdateJumplist", Sub() MainWindow.UpdateJumplist())
        RunIfQueued("RelaunchProgram", Sub() Application.Restart())
    End Sub

    Private Sub RunIfQueued(key As String, action As Action)
        If pendingActions.ContainsKey(key) AndAlso pendingActions(key) Then
            Try
                action.Invoke()
            Finally
                pendingActions(key) = False
            End Try
        End If
    End Sub

    Private Sub JumplistMaxItemsValBoxKeyPress(sender As Object, e As KeyPressEventArgs)
        If Control.ModifierKeys = Keys.None AndAlso Not Char.IsControl(e.KeyChar) Then
            JumplistMaxItemsTimer.Start()
        End If
    End Sub

    Private Sub JumplistMaxItemsValBoxValueChanged(sender As Object, e As EventArgs)
        JumplistMaxItemsTimer.Start()
    End Sub
    
    Private Sub JumplistMaxItemsTimerTick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        JumplistMaxItemsTimer.Stop()
        DidJumpListItemsMaximumChange()
    End Sub
    
    Public Sub DidJumpListItemsMaximumChange()
        If Globals.JumpListItemsMaximum <> JumplistMaxItemsValBox.Value Then
            Dim valInt = Convert.ToInt32(JumplistMaxItemsValBox.Value)
            Globals.JumpListItemsMaximum = valInt
            pendingActions("UpdateJumplist") = True

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "JumpListItems_Maximum", valInt, Microsoft.Win32.RegistryValueKind.DWord)
        End If
    End Sub

    Private Sub MetaItemsUpdateCheckboxToggled(sender As Object, e As EventArgs)
        Dim cb As CheckBox = DirectCast(sender, CheckBox)
        Dim newValue As Integer = If(cb.Checked, 1, 0) ' using CInt() makes True bool `-1` so this makes it `1` instead
        If newValue <> Globals.MetaSettings.HasUpdateItem Then
            Globals.MetaSettings.HasUpdateItem = newValue
            IniSettings.WriteValue("MetaUpdateItemEnabled", newValue)
            pendingActions("UpdateJumplist") = True
        End If
    End Sub

    Private Sub MetaItemsUpdateRefetch()
        Dim curFileValue = IniSettings.ReadValue("MetaUpdateItemEnabled")
        Dim curValue = Globals.MetaSettings.HasUpdateItem
        If curFileValue <> curValue Then
            Globals.MetaSettings.HasUpdateItem = curFileValue
            curValue = curFileValue
        End If
        Globals.MetaSettings.HasUpdateItem = curValue
        Me.MetaItemsUpdateCheckbox.Checked = curValue
    End Sub

    Private Sub LanguageDropdownMenuSelChanged(sender As Object, value As String)
        pendingActions("RelaunchProgram") = True
    End Sub

    Private Sub ThemeDropdownMenuSelChanged(sender As Object, value As String)
        pendingActions("RelaunchProgram") = True
    End Sub

    Private Sub MetaItemsUpdateShouldDisable()
        ' If meta items set to always hidden then disable update item checkbox from being selectable
        MetaItemsUpdateCheckbox.Enabled = (MetaItemsDropdownMenu.GetCurrentMode() <> 0)
    End Sub

    Private Sub MetaItemsDropdownMenuModeChanged(sender As Object, value As Integer)
        MetaItemsUpdateShouldDisable()
        pendingActions("UpdateJumplist") = True
    End Sub

    Private Sub DefaultLaunchActionMenuChanged(sender As Object, e As EventArgs)
        pendingActions("UpdateJumplist") = True
    End Sub
End Class
