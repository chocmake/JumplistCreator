Imports System.IO
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Linq
Imports System.Globalization

Public Class LangInfo
    Public Property Name As String
    Public Property Lang As String
End Class

Public Module Lang
    Private Const DEFAULT_LANG_STEM = "en-US"
    Private Dim currentLangStem As String
    Private ReadOnly embeddedLangPath As String =$"Languages.{DEFAULT_LANG_STEM}.txt"

    Private defaultLangDict As Dictionary(Of String, String)
    Private currentLangDict As Dictionary(Of String, String)
    Private languageDirectoryPath As String

    Sub New()
        defaultLangDict = LoadEmbeddedDefaultLanguage()
        languageDirectoryPath = DetermineLanguageDirectoryPath()

        Dim best As String = FindBestLanguageFile(Globalization.CultureInfo.CurrentUICulture.Name)
        SetLanguage(best)
    End Sub

    Public Sub InitLanguage()
        Dim iniLang = IniSettings.ReadValue("Language")
        If Not String.IsNullOrEmpty(iniLang) Then
            Lang.SetLanguage(iniLang)
        End If
    End Sub

    Public Sub SetLanguage(requestedCode As String)
        Dim actualCode = FindBestLanguageFile(requestedCode)
        Dim filePath = Path.Combine(languageDirectoryPath, $"{actualCode}.txt")

        If File.Exists(filePath) Then
            currentLangDict = LoadExternalLanguageFile(filePath)
            currentLangStem = actualCode
        Else
            currentLangDict = defaultLangDict
            currentLangStem = DEFAULT_LANG_STEM
        End If
    End Sub

    Private Function LoadEmbeddedDefaultLanguage() As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)
        Dim asm As Assembly = Assembly.GetExecutingAssembly()

        Using stream = asm.GetManifestResourceStream(embeddedLangPath)
            If stream IsNot Nothing Then
                Using reader As New StreamReader(stream)
                    Dim content = reader.ReadToEnd()
                    Dim lines = content.Split(Environment.NewLine.ToCharArray(),
                                              StringSplitOptions.RemoveEmptyEntries)

                    For Each line In lines
                        Dim parts = line.Split({"="c}, 2, StringSplitOptions.None)
                        If parts.Length = 2 Then
                            dict(parts(0).Trim()) = parts(1).Trim()
                        End If
                    Next
                End Using
            End If
        End Using

        Return dict
    End Function

   Public Function FindBestLanguageFile(cultureName As String) As String
       ' Try exact culture and its parent
       Dim ci As CultureInfo
       Try
           ci = CultureInfo.GetCultureInfo(cultureName)
       Catch ex As CultureNotFoundException
           ci = CultureInfo.InvariantCulture
       End Try

       Do
           If Not String.IsNullOrEmpty(ci.Name) Then
               Dim p = Path.Combine(languageDirectoryPath, ci.Name & ".txt")
               If File.Exists(p) Then
                   Return ci.Name
               End If
           End If
           ci = ci.Parent
       Loop While ci IsNot Nothing AndAlso Not String.IsNullOrEmpty(ci.Name)

       ' Sibling fallback for the same base (eg: for `fr-BE` check any `fr-*.txt`)
       Dim parts = cultureName.Split("-"c)
       Dim neutral = If(parts.Length > 0, parts(0).ToLowerInvariant(), String.Empty)

       If Not String.IsNullOrEmpty(neutral) _
          AndAlso Directory.Exists(languageDirectoryPath) Then

           Dim pattern = $"{neutral}-*.txt"
           Dim siblingFile = Directory.EnumerateFiles(languageDirectoryPath, pattern).FirstOrDefault()
           If siblingFile IsNot Nothing Then
               Return Path.GetFileNameWithoutExtension(siblingFile)
           End If
       End If

       ' Hard fallback
       Return "en"
   End Function

    Private Function DetermineLanguageDirectoryPath() As String
        Return Path.Combine(Path.GetDirectoryName(ExePath), "Languages")
    End Function

    Private Function LanguageFileExists(stem As String) As Boolean
        Dim filePath = Path.Combine(languageDirectoryPath, $"{stem}.txt")
        Return File.Exists(filePath)
    End Function

    Private Function GetLangPath(stem As String) As String
        Dim filePath = Path.Combine(languageDirectoryPath, $"{stem}.txt")
        If File.Exists(filePath)
            Return filePath
        End If
        Return Nothing
    End Function

    ' Unless no fallback specified uses the default dictionary as basis and masks over the key-value pairs from the external language file (to provide default fallbacks for any missing key-value pairs)
    Private Function LoadExternalLanguageFile(filePath As String, Optional noFallback As Boolean = False) As Dictionary(Of String, String)
        
        Dim dict As New Dictionary(Of String, String)
        If noFallback Then
            dict = New Dictionary(Of String, String)
        Else
            dict = New Dictionary(Of String, String)(defaultLangDict)
        End If

        Try
            Dim lines = File.ReadAllLines(filePath)
            For Each line In lines
                Dim parts = line.Split({"="c}, 2, StringSplitOptions.None)
                If parts.Length = 2 Then
                    dict(parts(0).Trim()) = parts(1).Trim()
                End If
            Next
        Catch ex As Exception
        End Try

        Return dict
    End Function

    Public Function GetCurrentLang() As String
        Return currentLangStem
    End Function

    Public Function GetString(key As String, Optional langStem As String = Nothing) As String
        If langStem IsNot Nothing Then
            Dim lookupLangDict As Dictionary(Of String, String)
            Dim lookupPath As String
            If langStem = DEFAULT_LANG_STEM Then
                lookupLangDict = defaultLangDict
            Else
                lookupPath = GetLangPath(langStem)
                If lookupPath IsNot Nothing Then
                    lookupLangDict = LoadExternalLanguageFile(lookupPath, noFallback:=True)
                Else
                    Return Nothing
                End If
            End If
            If lookupLangDict.ContainsKey(key) Then
                Return lookupLangDict(key)
            End If
            Return Nothing
        End If

        If currentLangDict.ContainsKey(key) Then
            Return currentLangDict(key)
        End If

        If defaultLangDict.ContainsKey(key) Then
            Return defaultLangDict(key)
        End If

        Return $"[Missing Translation: {key}]"
    End Function

    Public Function GetLangs() As Object()
        Dim files As New List(Of String)

        ' Check external files (if present)
        If Directory.Exists(languageDirectoryPath) Then
            files = Directory.GetFiles(languageDirectoryPath, "*.txt") _
                         .Select(Function(f) Path.GetFileNameWithoutExtension(f)) _
                         .ToList()
        End If

        ' Add embedded language
        If Not String.IsNullOrEmpty(DEFAULT_LANG_STEM) Then
            files.Add(DEFAULT_LANG_STEM)
        End If

        Dim result = files.Select(Function(token)
                    Dim display = GetString("LangDisplayName", token)
                    If display Is Nothing Then display = token
                    Return New LangInfo With {.Name = display, .Lang = token}
            End Function).ToArray()

        ' Sort by Lang key value
        Dim sorted As LangInfo() = DirectCast(result, LangInfo())
        Array.Sort(sorted, New KeyStringComparer(Of LangInfo)(Function(it) it.Lang))
        Return sorted
    End Function

End Module
