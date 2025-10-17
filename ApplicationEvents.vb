'
' Updated by chocmake (https://github.com/chocmake) from original, beginning 2025-07-20
'

Imports System.IO
Imports System.Globalization ' for locale override testing only
Imports System.Threading
Imports System.Reflection ' for reading assembly values

Namespace My
    Partial Friend Class MyApplication

        Private appMutex As Mutex

        Private Sub AppStart(sender As Object, e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup

            ' Set up single instance based on EXE path (so multiple different copies of the program can run but only one of each at a time)
            Dim hashId = FileUtilities.ComputeMd5(Globals.ExePath, False)
            Dim mutexName = $"Global\{Globals.ProgramName}_{hashId}"

            appMutex = New Mutex(False, mutexName)

            Dim obtainedInst As Boolean
            Try
                obtainedInst = appMutex.WaitOne(0, False)
            Catch ex As AbandonedMutexException
                obtainedInst = True
            End Try

            If Not obtainedInst Then
                e.Cancel = True
            End If

            ' Check for command-line arguments
            If e.CommandLine.Count > 0 Then
                Dim firstArg As String = e.CommandLine(0)
                Select Case firstArg.ToLowerInvariant()
                    Case "--config"
                        AssemblyResolver.Setup()

                    Case "--update"
                        AssemblyResolver.Setup()

                    Case Else
                        ' Check if drag-and-drop setting has been defined in settings for independent handling to the default launch action
                        Dim drop As ParsedJumplistEntry = CommandParser.ParseValueCommand(IniSettings.ReadValue("DefaultLaunchActionDrop"))
                        Dim dropTarget As String = drop.TargetPath
                        Dim dropArgs As String = drop.Args
                        If Not String.IsNullOrEmpty(dropTarget) AndAlso Not String.IsNullOrEmpty(dropArgs) Then
                            If dropTarget.Equals("<Default>", StringComparison.OrdinalIgnoreCase) Then
                                dropTarget = If(CommandParser.ParseValueCommand(IniSettings.ReadValue("DefaultLaunchAction")).TargetPath, dropTarget)
                            End If

                            ' Check if user has specified placeholder for first path arg or all path args (in case program only supports single paths or doesn't support multi-path syntax in space-delimited format)
                            If dropArgs.IndexOf("%1", StringComparison.Ordinal) >= 0 Then
                                dropArgs = dropArgs.Replace("%1", $"""{firstArg}""")
                            ElseIf dropArgs.IndexOf("%*", StringComparison.Ordinal) >= 0 Then
                                Dim joined As String = String.Join(" "c, e.CommandLine.Select(Function(a) """" & a & """"))
                                dropArgs = dropArgs.Replace("%*", joined)
                            End If

                            ' Call launch and if fails fall back to opening regular window
                            If Not CommandParser.LaunchCommand(dropTarget, dropArgs) Then
                                AssemblyResolver.Setup()
                            End If
                        Else
                            CommandParser.LaunchDefaultCommand()
                        End If

                End Select
            Else
                CommandParser.LaunchDefaultCommand()
            End If

            ' Testing only (locale override)
            ' Dim ci As New CultureInfo("es-ES")
            ' Thread.CurrentThread.CurrentCulture = ci
            ' Thread.CurrentThread.CurrentUICulture = ci
            ' Lang.SetLanguage(ci.Name)

            Lang.InitLanguage()
        End Sub

        Private Sub AppExit(sender As Object, e As EventArgs) Handles Me.Shutdown
            If appMutex IsNot Nothing Then
                Try
                    appMutex.ReleaseMutex()
                Catch
                End Try
                appMutex.Dispose()
                appMutex = Nothing
            End If
        End Sub
    End Class
End Namespace

Module Globals
    Private asm As Assembly = Assembly.GetEntryAssembly()
    
    Public ReadOnly Property ProgramName As String = asm.GetCustomAttribute(Of AssemblyTitleAttribute)()?.Title
    Public ReadOnly Property VersionNumber As String = asm.GetCustomAttribute(Of AssemblyFileVersionAttribute)()?.Version
    Public ReadOnly Property WebsiteUrl As String = "https://github.com/chocmake/JumplistCreator"
    
    Public Property ExePath As String = Application.ExecutablePath
    Public Property JumpListItemsMaximum As Integer = 0
    Public ReadOnly Property JumpListItemsMaximumDefault As Integer = 35
    Public Property MetaSettings As New MetaSettingValues()
    Public ReadOnly Property DefaultFont As Font = New Font("Segoe UI", 9.0F)

    Public Property DpiFactor As Single
    Public ReadOnly Property ComboBoxDefaultHeight As Integer = 22
    Public Const MARGIN_GENERIC As Integer = 12
End Module