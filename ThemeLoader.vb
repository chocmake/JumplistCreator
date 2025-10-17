Imports System.Reflection
Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Concurrent
Imports Microsoft.Win32

' Purpose of this module instead of a simpler namespace import is to conditionally load the DLL only if present beside the executable, to avoid needing it be present.

Public Module ThemeLoader
    Private ReadOnly dllPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Theme.dll")
    Private ReadOnly assemblyLock As New Object()
    Private loadedAssembly As Assembly = Nothing
    Private themeType As Type = Nothing
    Private ctorWithForm As ConstructorInfo = Nothing
    Private colorModeProp As PropertyInfo = Nothing

    ' Keep instances alive per-form
    Private ReadOnly instances As New ConcurrentDictionary(Of Form, Object)()

    Public Sub TryApplyTheme(target As Form)
        If Not DarkThemeIsActive() Then Return
        If target Is Nothing Then Return

        Try
            If Not EnsureTypeLoaded() Then Return
            If instances.ContainsKey(target) Then Return

            Dim ctor = ctorWithForm
            If ctor Is Nothing Then Return

            Dim instance As Object = ctor.Invoke(New Object() {target, False, True}) ' first bool = ColorizeIcons, second bool = RoundedPanels

            If colorModeProp IsNot Nothing Then
                Dim enumType = colorModeProp.PropertyType
                Dim enumVal As Object = Nothing
                Try
                    ' DarkModeForms supports `SystemDefault`, `ClearMode` and `DarkMode` but clear mode doesn't render controls fully like native so decided to only enable dark mode and handle conditional loading independently.
                    enumVal = [Enum].Parse(enumType, "DarkMode")
                Catch
                    Dim nested = themeType.GetNestedType("DisplayMode", BindingFlags.Public Or BindingFlags.NonPublic)
                    If nested IsNot Nothing Then
                        enumVal = [Enum].Parse(nested, "DarkMode")
                    End If
                End Try
                If enumVal IsNot Nothing Then
                    colorModeProp.SetValue(instance, enumVal)
                End If
            End If

            instances.TryAdd(target, instance)

            ' Remove entry when form closed
            AddHandler target.FormClosed, Sub(s, e) instances.TryRemove(target, Nothing)
        Catch ex As Exception
            MessageBox.Show(Lang.GetString("MsgThemeLoadError") & " " & ex.ToString(), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Function EnsureTypeLoaded() As Boolean
        If themeType IsNot Nothing Then Return True

        SyncLock assemblyLock
            If themeType IsNot Nothing Then Return True
            If Not ThemeLibExists() Then Return False

            Try
                loadedAssembly = Assembly.LoadFrom(dllPath)
                themeType = loadedAssembly.GetType("DarkModeForms.DarkModeCS", throwOnError:=False)
                If themeType Is Nothing Then Return False

                ctorWithForm = themeType.GetConstructor(New Type() {GetType(Form), GetType(Boolean), GetType(Boolean)})

                colorModeProp = themeType.GetProperty("ColorMode", BindingFlags.Public Or BindingFlags.Instance)
                Return True
            Catch
                themeType = Nothing
                loadedAssembly = Nothing
                ctorWithForm = Nothing
                colorModeProp = Nothing
                Return False
            End Try
        End SyncLock
    End Function

    Private Function IsUsingLightMode() As Boolean
        Const path As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const name As String = "AppsUseLightTheme"

        Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(path, writable:=False)
            If key Is Nothing Then
                Return True ' default to light theme
            End If

            Dim val As Object = key.GetValue(name)
            If val Is Nothing Then
                Return True
            End If

            Try
                Dim n As Integer = Convert.ToInt32(val)
                Return (n <> 0)
            Catch
                Return True
            End Try
        End Using
    End Function

    Public Function ThemeLibExists() As Boolean
        If File.Exists(dllPath) Then Return True
        Return False
    End Function

    Public Function DarkThemeIsActive() As Boolean
        ' Have to do it this way since any caller that needs to check what the dark theme state is needs to determine it early, prior to DarkModeForms subsequently applying the theme.
        Dim iniVal As String = IniSettings.ReadValue("Theme")
        Dim isDark As Boolean = False
        If iniVal = "Dark" AndAlso ThemeLibExists() Then
            isDark = True
        End If
        If iniVal = "Auto" Then
            If Not IsUsingLightMode() Then isDark = True
        End If
        Return isDark
    End Function
End Module
