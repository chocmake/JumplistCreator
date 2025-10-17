Imports System.IO
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.WindowsAPICodePack.Shell
Imports Microsoft.WindowsAPICodePack.Taskbar
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Imaging

Public Module IconExtractor
    Private Const SHGFI_ICON As Integer               = &H100 ' icon handle
    Private Const SHGFI_LARGEICON As Integer          = &H0 ' icon 32px or above
    Private Const SHGFI_SMALLICON As Integer          = &H1 ' 16x16px icon
    Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
    Private Const FILE_ATTRIBUTE_NORMAL As Integer    = &H80
    Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
    Private Const SHGFI_SYSICONINDEX As UInteger      = &H4000

    <StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto)> _
    Private Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst := 260)> _
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst := 80)> _
        Public szTypeName As String
    End Structure

    <DllImport("shell32.dll", CharSet := CharSet.Auto)> _
    Private Function SHGetFileInfo( _
        ByVal pszPath As String, _
        ByVal dwFileAttributes As Integer, _
        ByRef psfi As SHFILEINFO, _
        ByVal cbFileInfo As Integer, _
        ByVal uFlags As Integer _
    ) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError := True)> _
    Private Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

    Public Function GetShellIconIndex(inputPath As String, useFileAttributes As Boolean) As Integer
        Dim sfi As SHFILEINFO = New SHFILEINFO()
        Dim flags As UInteger = SHGFI_SYSICONINDEX Or SHGFI_ICON
        If useFileAttributes Then flags = flags Or SHGFI_USEFILEATTRIBUTES
        SHGetFileInfo(inputPath, 0, sfi, CUInt(Marshal.SizeOf(sfi)), flags)
        Return sfi.iIcon
    End Function

    ' Check if target has embedded icon
    Public Function HasLikelyEmbeddedIcon(inputPath As String) As Boolean
        If String.IsNullOrEmpty(inputPath) Then Return False

        ' Build fake 'associated' path to ask shell for the registered icon for this extension. If there's no extension use 'dummy' (generic file type)
        Dim ext = Path.GetExtension(inputPath)
        Dim assocPath As String
        If String.IsNullOrEmpty(ext) Then
            assocPath = "dummy"
        Else
            assocPath = "dummy" & ext
        End If

        Dim idxFile As Integer = GetShellIconIndex(inputPath, False) ' icon index returned by SHGetFileInfo (`False` to inspect the real filesystem object)
        Dim idxAssoc As Integer = GetShellIconIndex(assocPath, True) ' icon index for the fake associated path

        ' If the actual file's icon index differs from the associated/default icon index treat file as having likely embedded icon
        Return idxFile <> idxAssoc
    End Function

    Public Function GetShellIcon(ByVal inputPath As String, ByVal largeIcon As Boolean) As Icon
        Dim shfi As New SHFILEINFO()

        ' Specifically avoiding use of `SHGFI_USEFILEATTRIBUTES` flag since it only returns the generic dir icon, rather than any customized one
        Dim flags As Integer = SHGFI_ICON _
            Or (If(largeIcon, SHGFI_LARGEICON, SHGFI_SMALLICON))

        Dim attribs As Integer = If(IO.Directory.Exists(inputPath) Or inputPath.EndsWith("\"), _
            FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)

        Dim res As IntPtr = SHGetFileInfo(inputPath, attribs, shfi, Marshal.SizeOf(shfi), flags)
        If res = IntPtr.Zero Then
            Return Nothing
        End If

        ' Grab the HICON off the SHFILEINFO, turn it into a managed Icon then destroy the native handle
        Dim ico As Icon = Icon.FromHandle(shfi.hIcon).Clone()
        DestroyIcon(shfi.hIcon)
        Return ico
    End Function
End Module

Public Module IconParser
    Private Function GetLnkIcon(lnkPath As String) As IconReference
        Try
            If Not File.Exists(lnkPath) Then
                Return Nothing
            End If

            Dim shell As Object = CreateObject("WScript.Shell")
            Dim shortcut As Object = shell.CreateShortcut(lnkPath)
            Dim iconLocation As String = If(shortcut.IconLocation, String.Empty)
            Dim targetPath As String = If(shortcut.TargetPath, String.Empty)

            ' Directory icon special handling. There's no way afaict to only get the referenced file and index of directory icons, so instead the icon itself is extracted then saved to file and then that path referenced for the jumplist item.
            If Directory.Exists(targetPath) Then
                Return OutputIconFile(targetPath)
            End If

            ' If custom icon in LNK is defined (handle DLL/EXE/ICO paths)
            If Not String.IsNullOrEmpty(iconLocation) AndAlso iconLocation.Contains(","c) Then
                Dim parts = iconLocation.Split(","c)
                If parts.Length = 2 Then
                    Dim iconPath As String = parts(0).Trim().Trim(""""c)
                    Dim idxText As String = parts(1).Trim()
                    Dim idx As Integer = 0
                    Integer.TryParse(idxText, idx)

                    ' Expand and make absolute relative to the LNK folder
                    iconPath = Environment.ExpandEnvironmentVariables(iconPath)
                    Dim iconPathResolved As String = iconPath

                    ' Handling for targets that lack any icon path or embedded icon but which still contain an IconLocation (eg: batch script files)
                    If String.IsNullOrEmpty(iconPath) AndAlso Not IconExtractor.HasLikelyEmbeddedIcon(targetPath) Then
                        Return OutputIconFile(targetPath)
                    End If

                    If Not Path.IsPathRooted(iconPath) Then
                        iconPathResolved = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(lnkPath), iconPath))
                    End If

                    ' Try resolving against system file paths (eg: if `shell32.dll` by itself used for the icon path)
                    If Not File.Exists(iconPathResolved) Then
                        iconPathResolved = PathResolver.ResolveSystemFilePath(iconPath)
                    End If

                    If File.Exists(iconPathResolved) Then
                        Return New IconReference(iconPathResolved, idx)
                    End If
                End If
            End If

            ' Return icon from the direct target path
            If Not String.IsNullOrEmpty(targetPath) AndAlso File.Exists(targetPath) Then
                Return New IconReference(targetPath, 0)
            End If

            ' Otherwise fall back to the LNK file itself
            Return New IconReference(lnkPath, 0)

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function OutputIconFile(inputPath As String) As IconReference
        Dim appPath = My.Application.Info.DirectoryPath
        Dim icoDir = Path.Combine(appPath,"Icons")
        Dim icoLookup = Path.Combine(icoDir,"[Dedupe].csv")
        Dim icoSmall As Icon = IconExtractor.GetShellIcon(inputPath, False)
        Dim icoLarge As Icon = IconExtractor.GetShellIcon(inputPath, True)
        If icoSmall IsNot Nothing Then
            Try
                Directory.CreateDirectory(icoDir) ' create dir if it doesn't exist already
                Dim newFile = Path.Combine(icoDir, Guid.NewGuid().ToString("N") & ".ico")
                IconFileOutput.WithAlpha(newFile, icoSmall, icoLarge)
                Dim keptFilename = FileDeduplication.Dedupe(newFile, icoLookup)
                Return New IconReference(Path.Combine(icoDir, keptFilename), 0)
            Catch ex As Exception
                MessageBox.Show(Lang.GetString("MsgWriteFailureIcon") & Environment.NewLine & inputPath, $"{Globals.ProgramName}", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        Else
            Return New IconReference(inputPath, 0) ' fall back to using the path itself
        End If
    End Function

    Public Function GetIcon(inputPath As String) As IconReference
        ' First expand to absolute path in case environment vars used. Added specifically since StartAllBack's substitute taskbar has a bug that can't handle resolving icons for jumplists that contain environment variables.
        inputPath = Environment.ExpandEnvironmentVariables(inputPath)

        If Path.GetExtension(inputPath).ToLower() = ".lnk" Then
            Return GetLnkIcon(inputPath)
        End If

        ' Handle paths that lack embedded icons
        If Directory.Exists(inputPath) Or Not IconExtractor.HasLikelyEmbeddedIcon(inputPath) Then
            Return OutputIconFile(inputPath)
        End If

        ' Fall back to returning whatever icon is at that path
        Return New IconReference(inputPath, 0)
    End Function

    ' COM Interfaces for shell link operations
    <ComImport()>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("000214F9-0000-0000-C000-000000000046")>
    Private Interface IShellLink
        Sub GetPath(
            <Out(), MarshalAs(UnmanagedType.LPWStr)> pszFile As StringBuilder, 
            cchMaxPath As Integer, 
            <Out()> ByRef pfd As WIN32_FIND_DATA, 
            fFlags As UInteger)
    End Interface

    <ComImport()>
    <Guid("0000010B-0000-0000-C000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IPersistFile
        Sub Load(
            <MarshalAs(UnmanagedType.LPWStr)> pszFileName As String, 
            dwMode As Integer)
    End Interface

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WIN32_FIND_DATA
        Public dwFileAttributes As UInteger
        Public ftCreationTime As Long
        Public ftLastAccessTime As Long
        Public ftLastWriteTime As Long
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public cAlternateFileName As String
    End Structure

    <ComImport()>
    <ClassInterface(ClassInterfaceType.None)>
    <Guid("00021401-0000-0000-C000-000000000046")>
    Private Class ShellLinkClass
    End Class
End Module

Public Module IconFileOutput
    ' .NET's Icon.Save() only outputs to file with a limited color depth and no alpha channel, so this works around the limitation
    Public Sub WithAlpha(outputPath As String, smallIcon As Icon, Optional largeIcon As Icon = Nothing)
        If smallIcon Is Nothing Then Throw New ArgumentNullException(NameOf(smallIcon))

        ' If large size also passed and valid then combine both, otherwise just output small size
        Dim smallPng As Byte() = IconToPngBytes(smallIcon, 16, 16)
        Dim writeBoth As Boolean = (largeIcon IsNot Nothing)
        Dim largePng As Byte() = Nothing
        If writeBoth Then largePng = IconToPngBytes(largeIcon, 32, 32)

        Using fs As New FileStream(outputPath, FileMode.Create, FileAccess.Write)
            Using bw As New BinaryWriter(fs)
                bw.Write(CUShort(0))                    ' idReserved
                bw.Write(CUShort(1))                    ' idType = icon
                bw.Write(CUShort(If(writeBoth, 2, 1)))   ' idCount

                If writeBoth Then
                    Dim offsetFirst As Integer = 6 + (16 * 2)
                    Dim offsetSecond As Integer = offsetFirst + smallPng.Length

                    ' 16x16 size
                    WriteIconDirEntry(bw, 16, 16, 0, 0, 0, 0, CUInt(smallPng.Length), CUInt(offsetFirst))

                    ' 32x32 size
                    WriteIconDirEntry(bw, 32, 32, 0, 0, 0, 0, CUInt(largePng.Length), CUInt(offsetSecond))

                    bw.Write(smallPng)
                    bw.Write(largePng)
                Else
                    ' 16x16 size only
                    WriteIconDirEntry(bw, 16, 16, 0, 0, 0, 0, CUInt(smallPng.Length), CUInt(6 + 16))
                    bw.Write(smallPng)
                End If
            End Using
        End Using
    End Sub

    Private Sub WriteIconDirEntry(bw As BinaryWriter, bWidth As Integer, bHeight As Integer, bColorCount As Integer, bReserved As Integer, wPlanes As Integer, wBitCount As Integer, dwBytesInRes As UInteger, dwImageOffset As UInteger)
        bw.Write(CByte(If(bWidth >= 256, 0, bWidth)))   ' bWidth (0 means 256)
        bw.Write(CByte(If(bHeight >= 256, 0, bHeight))) ' bHeight
        bw.Write(CByte(bColorCount))
        bw.Write(CByte(bReserved))
        bw.Write(CUShort(wPlanes))
        bw.Write(CUShort(wBitCount))
        bw.Write(CUInt(dwBytesInRes))
        bw.Write(CUInt(dwImageOffset))
    End Sub

    Private Function IconToPngBytes(ic As Icon, width As Integer, height As Integer) As Byte()
        Using bmp As New Bitmap(width, height, PixelFormat.Format32bppArgb)
            Using g = Graphics.FromImage(bmp)
                g.Clear(Color.Transparent)
                g.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.HighQuality
                g.DrawIcon(ic, New Rectangle(0, 0, width, height))
            End Using
            Using ms As New MemoryStream()
                bmp.Save(ms, ImageFormat.Png)
                Return ms.ToArray()
            End Using
        End Using
    End Function
End Module

Module PathResolver
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function GetSystemDirectory(lpBuffer As System.Text.StringBuilder, uSize As Integer) As Integer
    End Function

    Private Function GetSystemDir() As String
        Dim sb As New System.Text.StringBuilder(260)
        If GetSystemDirectory(sb, sb.Capacity) > 0 Then
            Return sb.ToString()
        End If
        Return Environment.GetFolderPath(Environment.SpecialFolder.Windows)
    End Function

    Public Function ResolveSystemFilePath(input As String) As String
        If String.IsNullOrWhiteSpace(input) Then Return input

        Dim token = input.Trim().Trim(""""c)

        If Path.IsPathRooted(token) OrElse File.Exists(token) Then
            Return token
        End If

        Dim candidates As New List(Of String)

        Dim sys = GetSystemDir()
        If Not String.IsNullOrEmpty(sys) Then candidates.Add(Path.Combine(sys, token))

        Dim windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
        If Not String.IsNullOrEmpty(windir) Then candidates.Add(Path.Combine(windir, token))

        Dim pathEnv = Environment.GetEnvironmentVariable("PATH")
        If Not String.IsNullOrEmpty(pathEnv) Then
            For Each p In pathEnv.Split(";"c)
                If String.IsNullOrWhiteSpace(p) Then Continue For
                candidates.Add(Path.Combine(p.Trim(), token))
            Next
        End If

        For Each c In candidates
            Try
                If File.Exists(c) Then Return c
            Catch : End Try
        Next

        Return token
    End Function
End Module