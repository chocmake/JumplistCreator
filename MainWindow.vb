'
' Updated by chocmake (https://github.com/chocmake) from original `Form1.vb` beginning 2025-07-20
'

Imports System.IO
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.WindowsAPICodePack.Shell
Imports Microsoft.WindowsAPICodePack.Taskbar
Imports System.Runtime.InteropServices
Imports System.Text

Public Class MetaSettingValues
    Public Property Visibility As Integer = IniSettings.ReadValue("MetaItemsVisibility")
    Public Property HasUpdateItem As Integer = IniSettings.ReadValue("MetaUpdateItemEnabled")
    Public Property CategoryIndex As Integer
    Public Property ItemsCount As Integer
End Class

Partial Public Class MainWindow
    Dim args As String() = Environment.GetCommandLineArgs()

    ' Keep a reference to the Taskbar instance
    Private windowsTaskbar As TaskbarManager = TaskbarManager.Instance

    Private WithEvents UpdateAutoCloseTimer As New Timer
    Private currentValue As Integer = 0
    Private jumplistFile As String = FileUtilities.GetJumplistPath()

    Public Sub New()
        InitializeComponent()
        ThemeLoader.TryApplyTheme(Me)

        AddHandler Me.Load, AddressOf MainWindowLoad
        AddHandler Me.Closed, AddressOf MainWindowClosed
        AddHandler MyBase.Shown, AddressOf MainWindowShown
        AddHandler EditJumplistButton.Click, AddressOf EditJumplist
        AddHandler UpdateJumplistButton.Click, AddressOf UpdateJumplist
        AddHandler SettingsButton.Click, AddressOf OpenSettings
        AddHandler AboutButton.Click, AddressOf OpenAbout
        AddHandler UpdateAutoCloseTimer.Tick, AddressOf UpdateAutoCloseTimerTick
        AddHandler UpdateProgressTimer.Tick, AddressOf UpdateProgressTimerTick
    End Sub

    Private Sub MainWindowLoad(sender As Object, e As EventArgs)
        ' Check existing registry max value
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "JumpListItems_Maximum", Nothing) Is Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "JumpListItems_Maximum", Globals.JumpListItemsMaximumDefault, Microsoft.Win32.RegistryValueKind.DWord)
            Globals.JumpListItemsMaximum = Globals.JumpListItemsMaximumDefault
        Else
            Globals.JumpListItemsMaximum = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "JumpListItems_Maximum", Globals.JumpListItemsMaximumDefault)
        End If

        ' Configure form elements
        ReadWindowPos()

        ' Disabled state not currently applied prior to theme since buttons have centered text and DarkModeForms expects left-aligned for how it handles disabled control text
        If Not File.Exists(FileUtilities.GetJumplistPath()) Then
            EditJumplistButton.Text = Lang.GetString("GenerateJumplist")
            UpdateJumplistButton.Enabled = False
        End if
    End Sub

    Private Sub MainWindowShown(sender As Object, e As EventArgs)
        If Environment.GetCommandLineArgs().Count > 1 Then
            If Environment.GetCommandLineArgs(1) = "--update" Then
                StyleForUpdateArg()
                ' To update jumplists .NET requires an active program window to be open, hence why it's called on window shown
                If Not UpdateJumplist() Then
                    ' Handle if `--update` called but no jumplist INI exists
                    MessageBox.Show(Lang.GetString("MsgJumplistFileNotFound"), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        RestartAppToConfig() ' alternative to Application.Restart(), avoiding reusing the same launch args (which would otherwise cause loop)
                End If
            End If
        End If
    End Sub

    Private Sub RestartAppToConfig()
        Dim startInfo As System.Diagnostics.ProcessStartInfo = System.Diagnostics.Process.GetCurrentProcess().StartInfo
        startInfo.FileName = Globals.ExePath
        startInfo.Arguments = "--config"
        Dim exitMethod As System.Reflection.MethodInfo = GetType(Application).GetMethod("ExitInternal", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Static)
        If exitMethod IsNot Nothing Then
            exitMethod.Invoke(Nothing, Nothing)
        Else
            Application.Exit()
        End If
        System.Diagnostics.Process.Start(startInfo)
    End Sub

    Private Sub MainWindowClosed(sender As Object, e As EventArgs)
        SettingsWindow.DidJumpListItemsMaximumChange()
        WriteWindowPos()
    End Sub

    Private Sub ReadWindowPos()
        Dim pos As PointNullable = DirectCast(IniSettings.ReadValue("PriorWindowPos"), PointNullable)

        If pos.X.HasValue AndAlso pos.Y.HasValue Then
            ' Check that top-left of window from saved position is within screen bounds, otherwise fall back to default coords
            Dim r As Rectangle = Screen.GetBounds(New Point(pos.X.Value, pos.Y.Value))
            If r.Contains(New Point(pos.X.Value, pos.Y.Value)) Then
                Me.StartPosition = FormStartPosition.Manual
                Me.Location = New Point(pos.X.Value, pos.Y.Value)
            Else
                CenterOnCurrentScreen()
            End If
        Else
            CenterOnCurrentScreen()
        End If
    End Sub

    ' Substitute for `Me.StartPosition = FormStartPosition.CenterScreen` as since adding the adjustments for DPI awareness it hasn't accurately calculated coordinates
    Private Sub CenterOnCurrentScreen()
        Dim scr = Screen.FromPoint(Cursor.Position) ' treat active monitor as the one with cursor
        Dim area = scr.WorkingArea
        Dim x As Integer = area.Left + (area.Width - Me.Width) \ 2
        Dim y As Integer = area.Top + (area.Height - Me.Height) \ 2
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(x, y)
    End Sub

    Private Sub WriteWindowPos()
        Dim x As Integer = Me.Location.X
        Dim y As Integer = Me.Location.Y
        IniSettings.WriteValue("PriorWindowPos", $"{x},{y}")
    End Sub

    Private Sub EditJumplist(sender As Object, e As EventArgs)
        If File.Exists(jumplistFile) Then
            Process.Start(jumplistFile)
        Else
            Dim encoded As New UTF8Encoding(True) ' with BOM just for technically faster later INI encoding detection
            Try
                Using newFile As New StreamWriter(jumplistFile, True, encoded)
                    newFile.WriteLine("[✨ My Jumplist]")
                    newFile.WriteLine("Notepad=notepad.exe")
                    newFile.WriteLine("Pictures=%userprofile%\Pictures")
                    newFile.WriteLine("Recycle Bin=explorer.exe | args: shell:RecycleBinFolder | icon: shell32.dll,31")
                End Using
                Process.Start(jumplistFile)
                EditJumplistButton.Text = Lang.GetString("EditJumplist")
                UpdateJumplistButton.Enabled = True
                UpdateJumplist() ' update immediately just to show user how the example jumplist looks
            Catch ex As Exception
                MessageBox.Show(Lang.GetString("MsgWriteFailureIniJumplist"), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

    Public Function UpdateJumplist()
        If Not File.Exists(jumplistFile) Then Return False

        UpdateProgressBar.Value = 0
        UpdateProgressTimer.Start()

        ' Enable recent jumplists items in registry
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_TrackDocs", Nothing) Is Nothing Then _
            My.Computer.Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_TrackDocs")
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Start_TrackDocs", 1, Microsoft.Win32.RegistryValueKind.DWord)

        Dim jList As JumpList
        jList = JumpList.CreateJumpList()
        jList.ClearAllUserTasks()
        Dim category As JumpListCustomCategory = Nothing
        Dim openingSection As Boolean = False
        Dim link As JumpListLink = Nothing
        Dim iniLines = IniParser.Parse(jumplistFile)

        If Globals.MetaSettings.Visibility >= 1 Then
            Dim vis = MetaSettings.Visibility
            If (vis = 2) OrElse (vis = 1 AndAlso Not String.IsNullOrEmpty(IniSettings.ReadValue("DefaultLaunchAction"))) Then
                iniLines.Add(New IniSection With {
                    .Name  = "Jumplist"})
                iniLines.Add(New IniKeyValue With {
                    .Key   = Lang.GetString("MetaJumplistItemConfig"),
                    .Value = "<JumplistExe> | args: --config",
                    .IconPath = Globals.ExePath,
                    .IconIndex  = 1})
                If (Globals.MetaSettings.HasUpdateItem) Then
                    iniLines.Add(New IniKeyValue With {
                    .Key   = Lang.GetString("MetaJumplistItemUpdate"),
                    .Value = "<JumplistExe> | args: --update",
                    .IconPath = Globals.ExePath,
                    .IconIndex  = 2})
                End If
            End If
        End If

        Dim iniItemCountCur as Integer
        Dim iniItemCountTotal As Integer = iniLines.OfType(Of IniKeyValue)().Count()

        Dim lastSection = iniLines _
            .OfType(Of IniSection)() _
            .LastOrDefault()

        Dim keyValueCount As Integer = iniLines.OfType(Of IniKeyValue)().Count()
        Dim maxItemIndex As Integer = iniLines _
            .Select(Function(item, idx) New With {.Item = item, .Index = idx}) _
            .Where(Function(x) TypeOf x.Item Is IniKeyValue) _
            .Take(Globals.JumpListItemsMaximum + 1) _
            .Select(Function(x) x.Index) _
            .DefaultIfEmpty(-1) _
            .Max()
        Dim truncatedNonMetaIndex As Integer

        ' Check for the last meta named category and items
        If keyValueCount > Globals.JumpListItemsMaximum Then
            If lastSection IsNot Nothing AndAlso lastSection.Name = "Jumplist" Then
                Dim items = GetSectionItems(iniLines, "Jumplist", occurrence:=-1)

                If items.Count > 0 Then
                    ' The index of the last meta named category
                    Globals.MetaSettings.CategoryIndex = iniLines.FindLastIndex(
                        Function(el)
                            Dim sec = TryCast(el, IniSection)
                            Return sec IsNot Nothing _
                                AndAlso sec.Name.Equals("Jumplist", StringComparison.OrdinalIgnoreCase)
                        End Function)
                    Globals.MetaSettings.ItemsCount = iniLines _
                        .Skip(Globals.MetaSettings.CategoryIndex + 1) _
                        .TakeWhile(Function(item) TypeOf item Is IniKeyValue) _
                        .Count()
                    If Globals.MetaSettings.ItemsCount Then
                        truncatedNonMetaIndex = FindNthPriorKeyPairIndex(Of IniElement)(iniLines, maxItemIndex, Globals.MetaSettings.ItemsCount)
                    End If
                End If
            End If
        End If

        For i As Integer = 0 To iniLines.Count - 1
            If i >= truncatedNonMetaIndex AndAlso i < Globals.MetaSettings.CategoryIndex Then
                Continue For
            ' The `- 1` is to normalize for the zero-based counter
            ElseIf iniItemCountCur > (Globals.JumpListItemsMaximum - 1) Then
                ' Only truthy if no meta items present (to allow meta items to override max jumplist items value when max is lower value than the total number of meta items)
                If Globals.MetaSettings.ItemsCount = 0 Then
                    Continue For
                End if
            End If
            
            If TypeOf iniLines(i) Is IniMisc Then
                Continue For
            End If

            If TypeOf iniLines(i) Is IniSection Then
                Dim section = DirectCast(iniLines(i), IniSection)
                openingSection = True
                If category IsNot Nothing Then
                    jList.AddCustomCategories(category)
                End If
                category = New JumpListCustomCategory(section.Name)
                Continue For
            End If

            If TypeOf iniLines(i) Is IniKeyValue Then
                Dim item = DirectCast(iniLines(i), IniKeyValue)
                Dim result As ParsedJumplistEntry = CommandParser.ParseValueCommand(item.Value)
                Dim targetPath As String = result.TargetPath
                Dim arguments As String = result.Args
                Dim workingDir As String = result.StartIn
                Dim iconPath As String = FileUtilities.ResolvePath(result.IconPath)
                Dim iconIndex As Integer? = Math.Max(If(result.IconIndex, 0), 0) ' wrapped in max to avoid negative indices

                Dim AdjustCountersIndices = Sub()
                    ' Account for change in list counts/indices
                    keyValueCount -= 1
                    maxItemIndex -= 1
                    If truncatedNonMetaIndex > 0 Then
                        truncatedNonMetaIndex += 1
                    End If
                End Sub

                ' Validate path before adding as jumplist
                If CommandParser.ValidatePath(targetPath) Then
                    ' Check if path exists in current dir or PATH (for non-absolute path filenames)
                    Dim resolveResult = FileUtilities.ResolvePath(targetPath)
                    If resolveResult IsNot Nothing AndAlso resolveResult <> targetPath Then
                        targetPath = resolveResult
                    Else
                        ' Check if likely URL to prepend protocol if it's protocol-less (to avoid Windows Open With prompts) and add handler icon
                        Dim urlCheck = UrlParser.GetAbsoluteUri(targetPath)
                        If urlCheck.Item1 Then
                            targetPath = urlCheck.Item2.AbsoluteUri
                            item.IconPath = If(iconPath, UrlParser.GetSchemeHandlerPath(targetPath))
                            item.IconIndex = If(iconIndex, 0)
                        End If
                    End If

                    link = New JumpListLink(targetPath, item.Key) With {
                        .Arguments = arguments,
                        .IconReference =  If(Not String.IsNullOrEmpty(item.IconPath),
                                          New IconReference(item.IconPath, item.IconIndex),
                                          If(Not String.IsNullOrEmpty(iconPath),
                                             New IconReference(iconPath, iconIndex),
                                             IconParser.GetIcon(targetPath))),
                        .WorkingDirectory = workingDir
                    }

                    ' If this is the first key-value pair parsed and no jumplist category exists yet then create one
                    If Not openingSection Then
                        openingSection = True
                        category = New JumpListCustomCategory("Tasks")
                    End If
                    
                    category.AddJumpListItems(link)
                    iniItemCountCur += 1
                Else
                    AdjustCountersIndices()
                End If
                Continue For
            End If
        Next

        ' Add the last category
        If category IsNot Nothing Then
            jList.AddCustomCategories(category)
        End If

        jList.Refresh()

        If Environment.GetCommandLineArgs().Count > 1 Then
            If Environment.GetCommandLineArgs(1) = "--update" Then ' this will close form after update
                UpdateAutoCloseTimer.Interval = 800 ' delay 
                UpdateAutoCloseTimer.Start()
            End If
        End If

        Return True
    End Function

    Private Sub OpenSettings()
        Dim settingsDialog As New SettingsWindow()
        settingsDialog.StartPosition = FormStartPosition.CenterParent
        settingsDialog.ShowDialog(Me)
    End Sub

    Private Sub OpenAbout()
        Dim settingsDialog As New AboutWindow()
        settingsDialog.StartPosition = FormStartPosition.CenterParent
        settingsDialog.ShowDialog(Me)
    End Sub

    Private Function FindNthPriorKeyPairIndex(Of T)(items As IList(Of T), targetIndex As Integer, nth As Integer) As Integer
        If items Is Nothing Then Throw New ArgumentNullException(NameOf(items))
        If nth <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(nth))
        If targetIndex <= 0 Then Return -1
        If targetIndex > items.Count Then targetIndex = items.Count

        ' Work backward from the target index to find the nth prior IniKeyValue item
        Dim found As Integer = 0
        For i As Integer = targetIndex - 1 To 0 Step -1
            If TypeOf items(i) Is IniKeyValue Then
                found += 1
                If found = nth Then
                    Return i
                End If
            End If
        Next

        Return 0
    End Function

    Private Sub UpdateAutoCloseTimerTick(sender As Object, e As EventArgs)
        Application.Exit()
    End Sub

    Private Sub UpdateProgressTimerTick(sender As Object, e As EventArgs)
        Dim increment As Integer = 15
        If UpdateProgressBar.Value < UpdateProgressBar.Maximum Then
            ' Make sure non evenly divisible increments don't exceed max
            If UpdateProgressBar.Value + increment >= UpdateProgressBar.Maximum Then
                UpdateProgressBar.Value = UpdateProgressBar.Maximum
            Else
                UpdateProgressBar.Value += increment
            End If
        Else
            UpdateProgressTimer.Stop()
            UpdateProgressBar.Value = 0
        End If
    End Sub
End Class