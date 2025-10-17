Imports System.IO
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.Globalization
Imports System.Reflection

Public Class ParsedJumplistEntry
    Public Property TargetPath As String
    Public Property Args As String
    Public Property IconPath As String
    Public Property IconIndex As Integer?
    Public Property StartIn As String
End Class

Public Module AssemblyResolver
    Public Sub Setup()
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveAssemblies
    End Sub

    Private Function ResolveAssemblies(ByVal sender As Object, ByVal e As System.ResolveEventArgs) As Reflection.Assembly
        Dim desiredAssembly = New Reflection.AssemblyName(e.Name)
        If desiredAssembly.Name = "Microsoft.WindowsAPICodePack.Shell" Then
            Return Reflection.Assembly.Load(My.Resources.Microsoft_WindowsAPICodePack_Shell)
        ElseIf desiredAssembly.Name = "Microsoft.WindowsAPICodePack" Then
            Return Reflection.Assembly.Load(My.Resources.Microsoft_WindowsAPICodePack)
        Else
            Return Nothing
        End If
    End Function
End Module

Public Module FileUtilities
    <DllImport("kernel32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function SearchPath(lpPath As String, lpFileName As String, lpExtension As String, nBufferLength As Integer, lpBuffer As StringBuilder, ByRef lpFilePart As Integer) As Integer
    End Function

    Public Function GetIniPath(filename As String, Optional baseDir As String = Nothing) As String
        If String.IsNullOrEmpty(baseDir) Then
            ' If no reference path provided use executable's path for subsequent basis
            baseDir = Globals.ExePath
        End If

        ' This returns a path whether it exists yet or not, so on caller side check if path exists if necessary
        Return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(baseDir), filename)
    End Function

    Public Function GetJumplistPath(Optional baseDir As String = Nothing) As String
        Return GetIniPath("Jumplist.ini", baseDir)
    End Function

    Public Function GetSettingsPath(Optional baseDir As String = Nothing) As String
        Return GetIniPath("Settings.ini", baseDir)
    End Function

    Public Function ComputeMd5(input As String, Optional isFile As Boolean = False) As String
        Using md5 As MD5 = MD5.Create()
            Dim hashBytes As Byte()

            If isFile Then
                Using stream As FileStream = File.OpenRead(input)
                    hashBytes = md5.ComputeHash(stream)
                End Using
            Else
                Dim data As Byte() = Encoding.UTF8.GetBytes(input)
                hashBytes = md5.ComputeHash(data)
            End If

            ' Convert to hex string
            Dim sb As New StringBuilder(hashBytes.Length * 2)
            For Each b As Byte In hashBytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function

    ' Search PATH and current directory for input and return absolute path (similar but faster to `where` except doesn't match extensionless executables)
    Public Function ResolvePath(input As String, Optional workingDir As String = Nothing) As String
        If String.IsNullOrWhiteSpace(input) Then Return Nothing

        Dim initialSize As Integer = 260
        Dim sb As New StringBuilder(initialSize)
        Dim filePart As Integer = 0

        Dim ret = SearchPath(workingDir, input, Nothing, sb.Capacity, sb, filePart)
        If ret = 0 Then
            Return Nothing
        End If

        If ret > sb.Capacity Then
            ' Buffer was too small, resize and call again
            sb.Capacity = ret
            ret = SearchPath(workingDir, input, Nothing, sb.Capacity, sb, filePart)
            If ret = 0 Then Return Nothing
        End If

        Return sb.ToString()
    End Function

    Public Function LoadEmbeddedImage(name As String) As Image
        Dim asm = Assembly.GetExecutingAssembly()
        Using s = asm.GetManifestResourceStream(name)
            If s Is Nothing Then Return Nothing
            Return Image.FromStream(s)
        End Using
    End Function

    Public Function LoadEmbeddedIcon(name As String) As Icon
        Dim asm = Assembly.GetExecutingAssembly()
        Dim stream As Stream = asm.GetManifestResourceStream(name)
        If stream Is Nothing Then Return Nothing
        Try
            Return New Icon(stream)
        Finally
            stream.Dispose()
        End Try
    End Function
End Module

Public Module CommandParser
    <DllImport("user32.dll", SetLastError:=True)>
    Private Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    ' IShellLink will throw an unhandled exception if a path includes arguments accidentally by the user when added to a jumplist item. This can't be caught by wrapping the action in try/catch ime, so using WScript.Shell to pre-validate each item's path. This also prevents exceptions thrown from paths with illegal characters.
    Public Function ValidatePath(targetPath As String) As Boolean
        If String.IsNullOrWhiteSpace(targetPath) Then Return False

        Dim shell As Object = Nothing
        Dim shortcut As Object = Nothing
        Try
            shell = CreateObject("WScript.Shell")
            shortcut = shell.CreateShortcut(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") & ".lnk"))
            shortcut.TargetPath = targetPath
            Return True
        Catch ex As COMException
            Return False
        Catch ex As Exception
            Return False
        Finally
            Try
                If shortcut IsNot Nothing Then Marshal.ReleaseComObject(shortcut)
            Catch
            End Try
            Try
                If shell IsNot Nothing Then Marshal.ReleaseComObject(shell)
            Catch
            End Try
        End Try
    End Function

    ' Special handling for placeholder EXE paths for this program
    Private Function CheckMetaPlaceholder(val As String) As String
        If val.Equals("<JumplistExe>", StringComparison.OrdinalIgnoreCase) Then
            Return Globals.ExePath
        End If
        Return Unquote(val)
    End Function

    Public Function ParseValueCommand(ByVal line As String) As ParsedJumplistEntry
        If line Is Nothing Then
            Return Nothing
        End If

        Dim result As New ParsedJumplistEntry()

        ' Regex to find delimiters of the form: `| <optional space> <args|icon|dir>:`. This is to mitigate false positives if the pattern were instead to split on pipes alone since a value from an argument may include a pipe as some double quoted string.
        Dim delimRegex As New Regex("\|\s*(?=(args|icon|startin)\s*:)", RegexOptions.IgnoreCase)

        Dim matches As MatchCollection = delimRegex.Matches(line)

        If matches.Count = 0 Then
            ' If no delimiters then treat entire line as target path
            result.TargetPath = CheckMetaPlaceholder(line.Trim())
            Return result
        End If

        Dim firstMatch As Match = matches(0)
        Dim targetPart As String = line.Substring(0, firstMatch.Index).Trim()
        result.TargetPath = CheckMetaPlaceholder(targetPart)

        ' Iterate through delimiters to extract label:value slices
        ' For each match determine start of the label/value text (just after the `|`) and the end either as the next delimiter match index or the end of the string
        For i As Integer = 0 To matches.Count - 1
            Dim startIdx As Integer = matches(i).Index
            Dim afterPipeIdx As Integer = startIdx + 1

            Dim endIdx As Integer = If(i < matches.Count - 1, matches(i + 1).Index, line.Length)
            Dim segment As String = line.Substring(afterPipeIdx, endIdx - afterPipeIdx).Trim()

            Dim pairRegex As New Regex("^(?<label>args|icon|startin)\s*:\s*(?<value>.*)$", RegexOptions.IgnoreCase)
            Dim m As Match = pairRegex.Match(segment)
            If Not m.Success Then
                Continue For
            End If

            Dim label As String = m.Groups("label").Value.ToLowerInvariant()
            Dim value As String = Unquote(m.Groups("value").Value.Trim())

            Select Case label
                Case "args"
                    result.Args = value
                Case "icon"
                    ' Icon value can be either <path> or <path>,<index>
                    ' Since commas within paths need to preserved it's split on the last comma not inside quotes
                    Dim raw As String = m.Groups("value").Value.Trim()

                    ' If the raw value is of the form "quotedPath",index (quotes around path) then look for a leading quote. If present find the matching closing quote.
                    Dim iconPath As String = Nothing
                    Dim iconIndex As Integer? = Nothing

                    If raw.StartsWith("""") Then
                        Dim lastQuoteIdx As Integer = raw.LastIndexOf(""""c)
                        If lastQuoteIdx > 0 Then
                            Dim quotedPath As String = raw.Substring(1, lastQuoteIdx - 1)
                            ' Anything after the closing quote might be `,<index>`
                            Dim afterQuote As String = raw.Substring(lastQuoteIdx + 1).Trim()
                            If afterQuote.StartsWith(",") Then
                                Dim idxPart As String = afterQuote.Substring(1).Trim()
                                Dim parsedIdx As Integer = 0
                                If Integer.TryParse(idxPart, parsedIdx) Then
                                    iconPath = quotedPath
                                    iconIndex = parsedIdx
                                Else
                                    ' If no valid index after quoted path treat entire quoted content as path
                                    iconPath = quotedPath
                                    iconIndex = Nothing
                                End If
                            Else
                                ' No index present
                                iconPath = quotedPath
                                iconIndex = Nothing
                            End If
                        Else
                            ' Malformed quoting (single starting quote). Fall back to treating raw as path.
                            iconPath = Unquote(raw)
                            iconIndex = Nothing
                        End If
                    Else
                        ' Unquoted value. Split on last comma and attempt to parse trailing part as integer index.
                        Dim lastComma As Integer = raw.LastIndexOf(","c)
                        If lastComma >= 0 Then
                            Dim possiblePathPart As String = raw.Substring(0, lastComma).Trim()
                            Dim possibleIndexPart As String = raw.Substring(lastComma + 1).Trim()
                            Dim parsedIdx As Integer = 0
                            If Integer.TryParse(possibleIndexPart, parsedIdx) Then
                                iconPath = possiblePathPart
                                iconIndex = parsedIdx
                            Else
                                ' Trailing part not an integer. Treat whole raw as path.
                                iconPath = raw
                                iconIndex = Nothing
                            End If
                        Else
                            iconPath = raw
                            iconIndex = Nothing
                        End If
                    End If

                    ' Unquote iconPath if it happens to have surrounding quotes for whatever reason
                    If iconPath IsNot Nothing Then
                        iconPath = Unquote(iconPath)
                    End If

                    result.IconPath = iconPath
                    result.IconIndex = iconIndex
                Case "startin"
                    result.StartIn = value
            End Select
        Next

        Return result
    End Function

    ' Remove surrounding double quotes if both ends are quoted; preserves inner quotes.
    Private Function Unquote(ByVal s As String) As String
        If String.IsNullOrEmpty(s) Then
            Return s
        End If
        If s.Length >= 2 AndAlso s.StartsWith("""") AndAlso s.EndsWith("""") Then
            Return s.Substring(1, s.Length - 2)
        End If
        Return s
    End Function

    Public Function LaunchCommand(targetPath As String, arguments As String) As Boolean
        Try
            If String.IsNullOrEmpty(arguments) Then
                Dim psi As New ProcessStartInfo() With {
                    .FileName = targetPath,
                    .UseShellExecute = True}
                Dim p As Process = Process.Start(psi)
                SetForegroundWindow(p.MainWindowHandle) ' bring window to foreground
            Else
                Dim psi As New ProcessStartInfo() With {
                    .FileName = targetPath,
                    .Arguments = arguments,
                    .UseShellExecute = True}
                Dim p As Process = Process.Start(psi)
                SetForegroundWindow(p.MainWindowHandle)
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(Lang.GetString("MsgCustomLaunchActionFailure"), $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        Return False
    End Function

    Public Function LaunchDefaultCommand() As Boolean
        Dim value = IniSettings.ReadValue("DefaultLaunchAction")
        If Not String.IsNullOrEmpty(value) Then
            Dim result As ParsedJumplistEntry = ParseValueCommand(value)
            Dim targetPath As String = result.TargetPath
            Dim arguments As String = result.Args

            If Not String.IsNullOrEmpty(targetPath) Then
                If Not LaunchCommand(targetPath, arguments) Then
                    AssemblyResolver.Setup()
                    Return False
                Else
                    Return True
                End if
            End If
        End If

        AssemblyResolver.Setup() ' fallback if default unspecified
        Return False
    End Function
End Module

Public Module FileDeduplication
    Public Function Dedupe(
        newFileFullPath As String,
        indexCsvFullPath As String
    ) As String

        Dim newFileName = Path.GetFileName(newFileFullPath)
        If String.IsNullOrEmpty(newFileName) Then
            Throw New ArgumentException("Invalid file path.", NameOf(newFileFullPath))
        End If

        If Not File.Exists(newFileFullPath) Then
            Throw New FileNotFoundException("File not found.", newFileFullPath)
        End If

        Dim hash = FileUtilities.ComputeMd5(newFileFullPath, True)
        Dim index = LoadIndex(indexCsvFullPath)

        ' Check for an existing file with the same hash
        Dim duplicateFileName = index _
            .Where(Function(kv) kv.Value = hash) _
            .Select(Function(kv) kv.Key) _
            .FirstOrDefault()

        ' If a dupe exists delete the passed file
        If Not String.IsNullOrEmpty(duplicateFileName) Then
            File.Delete(newFileFullPath)
            Return duplicateFileName
        End If

        ' Otherwise store it in the index and return its filename
        AppendIndexEntry(indexCsvFullPath, newFileName, hash)
        Return newFileName

    End Function

    Private Function LoadIndex(csvPath As String) _
        As Dictionary(Of String, String)

        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        If Not File.Exists(csvPath) Then
            Return dict
        End If

        For Each line In File.ReadAllLines(csvPath)
            If String.IsNullOrWhiteSpace(line) Then Continue For
            Dim parts = line.Split(","c)
            If parts.Length <> 2 Then Continue For

            Dim fname = parts(0).Trim()
            Dim chk = parts(1).Trim()
            If Not dict.ContainsKey(fname) Then
                dict.Add(fname, chk)
            End If
        Next

        Return dict
    End Function

    Private Sub AppendIndexEntry(
        csvPath As String,
        fileName As String,
        checksum As String
    )
        Using w As New StreamWriter(csvPath, append:=True, encoding:=Encoding.UTF8)
            w.WriteLine($"{fileName},{checksum}")
        End Using
    End Sub

End Module

Public Module UrlParser
    Public Function GetAbsoluteUri(input As String) As Tuple(Of Boolean, Uri)
        Dim result As Uri = Nothing

        ' If lacks `<scheme>://` prepend `https://`
        If Not Regex.IsMatch(input, "^[A-Za-z][A-Za-z0-9+\-\.]*://", RegexOptions.IgnoreCase) Then
            input = "https://" & input
        End If

        If Not Uri.IsWellFormedUriString(input, UriKind.Absolute) Then
            Return Tuple.Create(False, CType(Nothing, Uri))
        End If

        If Not Uri.TryCreate(input, UriKind.Absolute, result) Then
            Return Tuple.Create(False, CType(Nothing, Uri))
        End If

        Dim ok As Boolean = (result IsNot Nothing AndAlso result.IsAbsoluteUri)
        Return Tuple.Create(ok, If(ok, result, CType(Nothing, Uri)))
    End Function

    Public Function GetSchemeHandlerPath(input As String) As String
        Dim uri As Uri = Nothing
        Uri.TryCreate(input, UriKind.Absolute, uri)
        If uri IsNot Nothing Then
            Dim handlerPath As String = AssocHelper.GetDefaultHandlerPathForScheme(uri.Scheme)
            If handlerPath isNot Nothing Then
                Return handlerPath
            End If
        End If
        Return Nothing
    End Function
End Module

Public Module AssocHelper
    Private Const ASSOCF_NONE As UInteger = 0
    Private Const ASSOCSTR_EXECUTABLE As UInteger = 2

    <DllImport("Shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function AssocQueryString(dwFlags As UInteger, str As UInteger, pszAssoc As String, pszExtra As String, pszOut As StringBuilder, ByRef pcchOut As UInteger) As Integer
    End Function

    Public Function GetDefaultHandlerPathForScheme(scheme As String) As String
        If String.IsNullOrWhiteSpace(scheme) Then Return Nothing
        Dim sb As New StringBuilder(260)
        Dim size As UInteger = CUInt(sb.Capacity)
        Dim hr As Integer = AssocQueryString(ASSOCF_NONE, ASSOCSTR_EXECUTABLE, scheme, Nothing, sb, size)
        If hr = 0 Then
            Return sb.ToString()
        End If
        ' If buffer was too small retry with required size
        If hr = &H8007007A Then ' ERROR_INSUFFICIENT_BUFFER
            sb.Capacity = CInt(size)
            hr = AssocQueryString(ASSOCF_NONE, ASSOCSTR_EXECUTABLE, scheme, Nothing, sb, size)
            If hr = 0 Then Return sb.ToString()
        End If
        Return Nothing
    End Function
End Module

' Sort an array by object key
' Eg:
    ' Dim sorted As MyType() = DirectCast(originalArray, MyType())
    ' Array.Sort(sorted, New KeyStringComparer(Of MyType)(Function(it) it.KeyToSortBy))
Public Class KeyStringComparer(Of T)
    Implements IComparer(Of T)

    Private ReadOnly _key As Func(Of T, String)

    Public Sub New(key As Func(Of T, String))
        _key = If(key, Function(x) If(x Is Nothing, String.Empty, x.ToString()))
    End Sub

    Public Function Compare(x As T, y As T) As Integer Implements IComparer(Of T).Compare
        Dim sx As String = _key(x)
        Dim sy As String = _key(y)
        Return String.Compare(sx, sy, StringComparison.CurrentCultureIgnoreCase)
    End Function
End Class

Public Module StringParser
    ' Truncate string based on how many visual glyphs fit in a given pixel space, to avoid differences in string truncation length between say Latin characters and wide Asian characters
    Public Function TruncateToWidth(text As String, font As Font, maxWidthPx As Single, trimFromEnd As Boolean) As String
        If String.IsNullOrEmpty(text) Then Return text

        Dim si As New StringInfo(text)
        Dim totalElements As Integer = si.LengthInTextElements

        Using bmp As New Bitmap(1, 1)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

                ' Check whole string fits
                Dim sizeAll As SizeF = g.MeasureString(text, font)
                If sizeAll.Width <= maxWidthPx Then Return text

                Dim lo As Integer = 1
                Dim hi As Integer = totalElements
                Dim best As String = ""

                While lo <= hi
                    Dim mid As Integer = (lo + hi) \ 2
                    Dim candidate As String

                    If trimFromEnd Then
                        candidate = si.SubstringByTextElements(totalElements - mid, mid)
                    Else
                        candidate = si.SubstringByTextElements(0, mid)
                    End If

                    Dim sz As SizeF = g.MeasureString(candidate, font)

                    If sz.Width <= maxWidthPx Then
                        best = candidate
                        lo = mid + 1
                    Else
                        hi = mid - 1
                    End If
                End While

                Return best
            End Using
        End Using
    End Function
End Module

Module DpiParser
    <DllImport("user32.dll", EntryPoint:="GetDpiForWindow", SetLastError:=False)>
    Private Function GetDpiForWindow(formHandle As IntPtr) As UInteger
    End Function

    Public Function GetFactor(form As Form) As Single
        Try
            Dim formHandle As IntPtr = form.Handle
            Dim dpi As UInteger = GetDpiForWindow(formHandle)
            If dpi = 0 Then Return 1.0F
            Return dpi / 96.0F
        Catch ex As EntryPointNotFoundException
            Return 1.0F
        End Try
    End Function
End Module