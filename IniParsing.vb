Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Linq
Imports System.Collections.Generic

Public MustInherit Class IniElement
End Class

Public Class IniSection
    Inherits IniElement
    Public Property Name As String
End Class

Public Class IniKeyValue
    Inherits IniElement
    Public Property Key As String
    Public Property Value As String
    Public Property IconPath As String
    Public Property IconIndex As Integer
End Class

' Comment/empty/invalid line
Public Class IniMisc
    Inherits IniElement
    Public Property RawLine As String
End Class

Public Class PointNullable
    Public Property X As Integer?
    Public Property Y As Integer?
    Public Sub New(Optional x As Integer? = Nothing, Optional y As Integer? = Nothing)
        Me.X = x
        Me.Y = y
    End Sub
End Class

Public Module IniParser

    Private SectionRegex As New Regex("^\s*\[(.+?)\]\s*$")
    Private KeyValueRegex As New Regex("^\s*([^=;]+?)\s*=\s*(.*?)\s*$")

    ' Parse in sequential order, to allow key-value lines that lack sections to be readable back in the same order
    Public Function Parse(filePath As String) As List(Of IniElement)
        If Not File.Exists(filePath) Then Return Nothing

        Dim lines = File.ReadAllLines(filePath, DetectEncoding(filePath))
        Dim result As New List(Of IniElement)(lines.Length)

        For Each raw In lines
            Dim line = raw.Trim()

            ' Empty/comment lines
            If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") Then
                result.Add(New IniMisc With {.RawLine = raw})
                Continue For
            End If

            Dim matchSection = SectionRegex.Match(line)
            If matchSection.Success Then
                result.Add(New IniSection With {.Name = matchSection.Groups(1).Value})
                Continue For
            End If

            Dim matchKv = KeyValueRegex.Match(line)
            If matchKv.Success Then
                result.Add(New IniKeyValue With {
                    .Key   = matchKv.Groups(1).Value.Trim(),
                    .Value = matchKv.Groups(2).Value.Trim()})
                Continue For
            End If

            ' Anything else
            result.Add(New IniMisc With {.RawLine = raw})
        Next

        Return result
    End Function

    Public Function GetSectionItems(
        elements As List(Of IniElement),
        sectionName As String,
        Optional occurrence As Integer = -1
    ) As List(Of IniKeyValue)
        Dim matches As New List(Of List(Of IniKeyValue))
        Dim current As List(Of IniKeyValue) = Nothing

        For Each el As IniElement In elements
            If TypeOf el Is IniSection Then
                Dim section = DirectCast(el, IniSection)
                If String.Equals(section.Name, sectionName, StringComparison.Ordinal) Then
                    ' Start a new section block for this occurrence of that section
                    current = New List(Of IniKeyValue)
                    matches.Add(current)
                Else
                    current = Nothing
                End If

            ElseIf TypeOf el Is IniKeyValue Then
                Dim kv = DirectCast(el, IniKeyValue)
                ' If within the last opened section block append its key-value
                If current IsNot Nothing Then
                    current.Add(kv)
                End If
            End If
        Next

        If matches.Count = 0 Then
            Return New List(Of IniKeyValue)()
        End If

        If occurrence = -1 OrElse occurrence >= matches.Count Then
            ' Return last occurrence
            Return matches(matches.Count - 1)
        Else
            Return matches(occurrence)
        End If
    End Function

    Public Function ReadValue(
        filePath As String,
        sectionName As String,
        keyName As String,
        Optional defaultValue As String = ""
    ) As String
        Dim elements = Parse(filePath)

        ' Grab key-value list from the last occurrence
        Dim kvList As List(Of IniKeyValue) = GetSectionItems(elements, sectionName, occurrence:=-1)

        ' Find the first matching key
        Dim found = kvList _
            .FirstOrDefault(Function(kv) String.Equals(kv.Key, keyName, StringComparison.Ordinal))

        If found Is Nothing OrElse String.IsNullOrEmpty(found.Key) Then
            Return defaultValue
        End If

        Return found.Value
    End Function

    Public Sub WriteValue(
        filePath As String,
        section As String,
        keyName As String,
        newValue As String
    )
        Dim encoding = DetectEncoding(filePath)

        ' Read all lines (or start empty)
        Dim lines As List(Of String) =
            If(File.Exists(filePath),
               File.ReadAllLines(filePath, encoding).ToList(),
               New List(Of String)())

        Dim currentSect As String = String.Empty
        Dim sectLineIndex As Integer = -1
        Dim keyLineIndex As Integer = -1

        ' Locate the section header and key
        For i = 0 To lines.Count - 1
            Dim raw = lines(i)
            Dim line = raw.Trim()

            ' Detect section
            Dim matchSect = SectionRegex.Match(line)
            If matchSect.Success Then
                currentSect = matchSect.Groups(1).Value
                If currentSect.Equals(section, StringComparison.OrdinalIgnoreCase) Then
                    sectLineIndex = i
                End If
                Continue For
            End If

            ' Detect key under the found section
            If sectLineIndex >= 0 AndAlso currentSect.Equals(section, StringComparison.OrdinalIgnoreCase) Then
                Dim matchKv = KeyValueRegex.Match(line)
                If matchKv.Success AndAlso matchKv.Groups(1).Value.Equals(keyName, StringComparison.OrdinalIgnoreCase) Then
                    keyLineIndex = i
                    Exit For
                End If
            End If
        Next

        ' If section missing append it
        If sectLineIndex < 0 Then
            ' Check if empty file and skip adding empty new line if so
            If lines.Count > 0 AndAlso lines.Any(Function(s) Not String.IsNullOrWhiteSpace(s)) Then
                lines.Add(String.Empty)
            End If
            lines.Add("[" & section & "]")
            sectLineIndex = lines.Count - 1
        End If

        Dim newLine = keyName & "=" & newValue

        If keyLineIndex >= 0 Then
            lines(keyLineIndex) = newLine ' replace existing
        Else
            lines.Insert(sectLineIndex + 1, newLine) ' insert below section header
        End If

        Try
            File.WriteAllLines(filePath, lines, encoding)
        Catch ex As Exception
            MessageBox.Show(Lang.GetString("MsgWriteFailureIniGeneric"), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Function DetectEncoding(path As String) As Encoding
        Using fs = New FileStream(path, FileMode.Open, FileAccess.Read)
            If fs.Length >= 3 Then
                Dim bom(3) As Byte
                fs.Read(bom, 0, 4)
                ' UTF-8 BOM: EF BB BF
                If bom(0) = &HE AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                    Return New UTF8Encoding(True)
                End If
                ' UTF-16 LE BOM: FF FE
                If bom(0) = &HFF AndAlso bom(1) = &HFE Then
                    Return New UnicodeEncoding(False, True)
                End If
                ' UTF-16 BE BOM: FE FF
                If bom(0) = &HFE AndAlso bom(1) = &HFF Then
                    Return New UnicodeEncoding(True, True)
                End If
            End If
        End Using

        ' If no BOM treat as UTF-8
        Return New UTF8Encoding(False)
   End Function

End Module

Public Module IniSettings
    Public Delegate Function ReadDelegate() As Object

    ' MetaItemsVisibility: `0` for hidden, `1` for conditionally visible (only when DefaultLaunchAction set) and `2` for always visible
    Private ReadOnly lookup As New Dictionary(Of String, ReadDelegate) From {
        {"PriorWindowPos", Function() CType(ParseCoordsFromString(), Object)},
        {"Theme", Function() IniSettings.ReadValueDirect("Theme", allowedValues:=New String() {"Auto", "Dark", "Light"}, defaultValue:="Auto")},
        {"MetaItemsVisibility", Function() Integer.Parse(IniSettings.ReadValueDirect("MetaItemsVisibility", allowedValues:=New String() {"0", "1", "2"}, defaultValue:="1"))},
        {"MetaUpdateItemEnabled", Function() Integer.Parse(IniSettings.ReadValueDirect("MetaUpdateItemEnabled", allowedValues:=New String() {"0", "1"}, defaultValue:="0"))}
    }

    Public Function ReadValue(key As String) As Object
        Dim d As ReadDelegate = Nothing
        If lookup.TryGetValue(key, d) Then
            Return d.Invoke()
        End If

        ' If no special handling just return direct INI lookup
        Return ReadValueDirect(key)
    End Function

    Public Function ReadValueDirect(
        keyName As String,
        Optional allowedValues As IEnumerable(Of String) = Nothing,
        Optional defaultValue As String = ""
    ) As String
        Dim filePath as String = FileUtilities.GetSettingsPath()
        If Not File.Exists(filePath) Then Return defaultValue
        Dim sectionName as String = "Settings"

        Dim raw = IniParser.ReadValue(filePath, sectionName, keyName).Trim()
        If String.IsNullOrEmpty(raw) Then
            Return defaultValue
        End If

        If allowedValues IsNot Nothing Then
            ' Case-sensitive comparison
            If allowedValues.Any(Function(v) String.Equals(v, raw, StringComparison.Ordinal)) Then
                Return raw
            Else
                Return defaultValue
            End If
        End If

        Return raw
    End Function

    Public Sub WriteValue(
        keyName As String,
        Optional value As String = "" ' set empty value if none passed
    )
        Dim filePath as String = FileUtilities.GetSettingsPath()
        If Not File.Exists(filePath) Then
            ' Create file
            Dim encoded As New UTF8Encoding(True)
            Try
                Using newFile As New StreamWriter(filePath, True, encoded)
                End Using
            Catch ex As Exception
                MessageBox.Show(Lang.GetString("MsgWriteFailureIniSettings"), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
        Dim sectionName as String = "Settings"

        Try
            IniParser.WriteValue(filePath, sectionName, keyName, value)
        Catch ex As Exception
        End Try
    End Sub

    Private Function ParseCoordsFromString() As PointNullable
        Dim iniVal As String = IniSettings.ReadValueDirect("PriorWindowPos")
        Dim p As New PointNullable()
        If String.IsNullOrWhiteSpace(iniVal) Then Return p
        Dim parts = iniVal.Split(","c)
        Dim tmp As Integer
        If parts.Length > 0 AndAlso Integer.TryParse(parts(0).Trim(), tmp) Then p.X = tmp
        If parts.Length > 1 AndAlso Integer.TryParse(parts(1).Trim(), tmp) Then p.Y = tmp
        Return p
    End Function
End Module