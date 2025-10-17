Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms

Public Class ImageMaskButton
    Inherits Control

    Public Enum ImageState
        Normal = 0
        Hover = 1
        Pressed = 2
        Focused = 3
        Disabled = 4
    End Enum

    Private imgState As ImageState = ImageState.Normal
    Private mouseDownInside As Boolean = False

    ' Expects white color PNG with alpha as white color will be used for tinting
    Private maskImageVal As Image
    Private maskHoverVal As Image
    Private maskPressedVal As Image
    Private maskFocusedVal As Image
    Private maskDisabledVal As Image

    <Browsable(True), Category("Appearance")>
    Public Property MaskImage As Image
        Get
            Return maskImageVal
        End Get
        Set(value As Image)
            If maskImageVal IsNot value Then
                maskImageVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property MaskHover As Image
        Get
            Return maskHoverVal
        End Get
        Set(value As Image)
            If maskHoverVal IsNot value Then
                maskHoverVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property MaskPressed As Image
        Get
            Return maskPressedVal
        End Get
        Set(value As Image)
            If maskPressedVal IsNot value Then
                maskPressedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property MaskFocused As Image
        Get
            Return maskFocusedVal
        End Get
        Set(value As Image)
            If maskFocusedVal IsNot value Then
                maskFocusedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property MaskDisabled As Image
        Get
            Return maskDisabledVal
        End Get
        Set(value As Image)
            If maskDisabledVal IsNot value Then
                maskDisabledVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    ' Color/opacity per state
    Private colorNormalVal As Color = Color.Black
    Private colorHoverVal As Color = Color.Black
    Private colorPressedVal As Color = Color.Black
    Private colorFocusedVal As Color = Color.Black
    Private colorDisabledVal As Color = Color.Gray

    Private alphaNormalVal As Single = 1.0F
    Private alphaHoverVal As Single = 1.0F
    Private alphaPressedVal As Single = 1.0F
    Private alphaFocusedVal As Single = 1.0F
    Private alphaDisabledVal As Single = 0.5F

    <Browsable(True), Category("Appearance")>
    Public Property ColorNormal As Color
        Get
            Return colorNormalVal
        End Get
        Set(value As Color)
            If colorNormalVal <> value Then
                colorNormalVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property ColorHover As Color
        Get
            Return colorHoverVal
        End Get
        Set(value As Color)
            If colorHoverVal <> value Then
                colorHoverVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property ColorPressed As Color
        Get
            Return colorPressedVal
        End Get
        Set(value As Color)
            If colorPressedVal <> value Then
                colorPressedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property ColorDisabled As Color
        Get
            Return colorDisabledVal
        End Get
        Set(value As Color)
            If colorDisabledVal <> value Then
                colorDisabledVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property ColorFocused As Color
        Get
            Return colorFocusedVal
        End Get
        Set(value As Color)
            If colorFocusedVal <> value Then
                colorFocusedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property AlphaNormal As Single
        Get
            Return alphaNormalVal
        End Get
        Set(value As Single)
            value = Math.Max(0F, Math.Min(1F, value))
            If Math.Abs(alphaNormalVal - value) > 0.001F Then
                alphaNormalVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property AlphaHover As Single
        Get
            Return alphaHoverVal
        End Get
        Set(value As Single)
            value = Math.Max(0F, Math.Min(1F, value))
            If Math.Abs(alphaHoverVal - value) > 0.001F Then
                alphaHoverVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property AlphaPressed As Single
        Get
            Return alphaPressedVal
        End Get
        Set(value As Single)
            value = Math.Max(0F, Math.Min(1F, value))
            If Math.Abs(alphaPressedVal - value) > 0.001F Then
                alphaPressedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property AlphaDisabled As Single
        Get
            Return alphaDisabledVal
        End Get
        Set(value As Single)
            value = Math.Max(0F, Math.Min(1F, value))
            If Math.Abs(alphaDisabledVal - value) > 0.001F Then
                alphaDisabledVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    <Browsable(True), Category("Appearance")>
    Public Property AlphaFocused As Single
        Get
            Return alphaFocusedVal
        End Get
        Set(value As Single)
            value = Math.Max(0F, Math.Min(1F, value))
            If Math.Abs(alphaFocusedVal - value) > 0.001F Then
                alphaFocusedVal = value
                InvalidateCache()
                Invalidate()
            End If
        End Set
    End Property

    ' Cached tinted bitmaps per state
    Private cachedNormalVal As Bitmap = Nothing
    Private cachedHoverVal As Bitmap = Nothing
    Private cachedPressedVal As Bitmap = Nothing
    Private cachedFocusedVal As Bitmap = Nothing
    Private cachedDisabledVal As Bitmap = Nothing
    Private cachedSizeVal As Size = Size.Empty

    Public Sub New()
        DoubleBuffered = True
        TabStop = True
        Size = New Size(100, 30)
    End Sub

    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        InvalidateCache()
        Invalidate()
    End Sub

    Private Sub InvalidateCache()
        cachedNormalVal?.Dispose() : cachedNormalVal = Nothing
        cachedHoverVal?.Dispose() : cachedHoverVal = Nothing
        cachedPressedVal?.Dispose() : cachedPressedVal = Nothing
        cachedFocusedVal?.Dispose() : cachedFocusedVal = Nothing
        cachedDisabledVal?.Dispose() : cachedDisabledVal = Nothing
        cachedSizeVal = Size.Empty
    End Sub

    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        imgState = If(Enabled, ImageState.Normal, ImageState.Disabled)
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        If Enabled Then imgState = ImageState.Hover : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        If Enabled Then
            mouseDownInside = False
            imgState = If(Focused, ImageState.Focused, ImageState.Normal)
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If Enabled AndAlso e.Button = MouseButtons.Left Then
            mouseDownInside = True
            imgState = ImageState.Pressed
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If Enabled AndAlso e.Button = MouseButtons.Left Then
            If mouseDownInside AndAlso ClientRectangle.Contains(e.Location) Then OnClick(EventArgs.Empty)
            mouseDownInside = False
            imgState = If(ClientRectangle.Contains(PointToClient(MousePosition)), ImageState.Hover, If(Focused, ImageState.Focused, ImageState.Normal))
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If Enabled AndAlso (e.KeyCode = Keys.Space OrElse e.KeyCode = Keys.Enter) Then
            imgState = ImageState.Pressed
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
        If Enabled AndAlso (e.KeyCode = Keys.Space OrElse e.KeyCode = Keys.Enter) Then
            imgState = If(Focused, ImageState.Focused, ImageState.Normal)
            OnClick(EventArgs.Empty)
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        If Enabled Then
            imgState = If(ClientRectangle.Contains(PointToClient(MousePosition)), ImageState.Hover, ImageState.Focused)
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e)
        If Enabled Then
            imgState = If(ClientRectangle.Contains(PointToClient(MousePosition)), ImageState.Hover, ImageState.Normal)
            Invalidate()
        End If
    End Sub

    Private Function EnsureCachedBitmaps() As Boolean
        If maskImageVal Is Nothing Then Return False
        If cachedSizeVal = Me.ClientSize AndAlso cachedNormalVal IsNot Nothing Then Return True

        InvalidateCache() ' dispose old

        cachedSizeVal = Me.ClientSize
        
        ' Obtain a resized mask bitmap for a given source image
        Dim getResizedMask As Func(Of Image, Bitmap) =
            Function(src As Image) As Bitmap
                If src Is Nothing Then Return Nothing
                Dim b As New Bitmap(Width, Height, PixelFormat.Format32bppArgb)
                Using g = Graphics.FromImage(b)
                    g.Clear(Color.Transparent)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(src, New Rectangle(0, 0, Width, Height),
                                0, 0, src.Width, src.Height, GraphicsUnit.Pixel)
                End Using
                Return b
            End Function

        ' Create masks per state (if none specified fall back to primary MaskImage)
        Dim maskNormalBmp = getResizedMask(maskImageVal)
        Dim maskHoverBmp = If(maskHoverVal IsNot Nothing, getResizedMask(maskHoverVal), maskNormalBmp)
        Dim maskPressedBmp = If(maskPressedVal IsNot Nothing, getResizedMask(maskPressedVal), maskNormalBmp)
        Dim maskFocusedBmp = If(maskFocusedVal IsNot Nothing, getResizedMask(maskFocusedVal), maskNormalBmp)
        Dim maskDisabledBmp = If(maskDisabledVal IsNot Nothing, getResizedMask(maskDisabledVal), maskNormalBmp)

        ' Create tinted bitmap from the given resized mask bitmap
        Dim createFromMask As Func(Of Bitmap, Color, Single, Bitmap) =
            Function(maskBmpLocal As Bitmap, col As Color, alpha As Single) As Bitmap
                If maskBmpLocal Is Nothing Then Return Nothing
                Dim out As New Bitmap(Width, Height, PixelFormat.Format32bppArgb)
                Using gOut = Graphics.FromImage(out)
                    gOut.Clear(Color.Transparent)
                    Dim r = col.R / 255.0F
                    Dim gC = col.G / 255.0F
                    Dim b = col.B / 255.0F
                    Dim a = Math.Max(0F, Math.Min(1F, alpha))
                    Dim cm As New ColorMatrix(New Single()() {
                        New Single() {r, 0, 0, 0, 0},
                        New Single() {0, gC, 0, 0, 0},
                        New Single() {0, 0, b, 0, 0},
                        New Single() {0, 0, 0, a, 0},
                        New Single() {0, 0, 0, 0, 1}
                    })
                    Using ia As New ImageAttributes()
                        ia.SetColorMatrix(cm, ColorMatrixFlag.Default, ColorAdjustType.Bitmap)
                        ia.SetWrapMode(Drawing2D.WrapMode.TileFlipXY)
                        gOut.DrawImage(maskBmpLocal, New Rectangle(0, 0, Width, Height),
                                       0, 0, maskBmpLocal.Width, maskBmpLocal.Height, GraphicsUnit.Pixel, ia)
                    End Using
                End Using
                Return out
            End Function

        ' Create cached bitmaps using per-state masks
        cachedNormalVal = createFromMask(maskNormalBmp, colorNormalVal, alphaNormalVal)
        cachedHoverVal = createFromMask(maskHoverBmp, colorHoverVal, alphaHoverVal)
        cachedPressedVal = createFromMask(maskPressedBmp, colorPressedVal, alphaPressedVal)
        cachedFocusedVal = createFromMask(maskFocusedBmp, colorFocusedVal, alphaFocusedVal)
        cachedDisabledVal = createFromMask(maskDisabledBmp, colorDisabledVal, alphaDisabledVal)

        ' Dispose intermediate resized masks (but don't dispose maskNormalBmp if reused by others)
        If maskNormalBmp IsNot Nothing Then maskNormalBmp.Dispose()
        If maskHoverBmp IsNot Nothing AndAlso maskHoverBmp IsNot maskNormalBmp Then maskHoverBmp.Dispose()
        If maskPressedBmp IsNot Nothing AndAlso maskPressedBmp IsNot maskNormalBmp Then maskPressedBmp.Dispose()
        If maskFocusedBmp IsNot Nothing AndAlso maskFocusedBmp IsNot maskNormalBmp Then maskFocusedBmp.Dispose()
        If maskDisabledBmp IsNot Nothing AndAlso maskDisabledBmp IsNot maskNormalBmp Then maskDisabledBmp.Dispose()

        Return True
    End Function

    Private Function GetCachedBitmapForState(s As ImageState) As Bitmap
        If Not EnsureCachedBitmaps() Then Return Nothing
        Select Case s
            Case ImageState.Normal
                Return cachedNormalVal
            Case ImageState.Hover
                Return cachedHoverVal
            Case ImageState.Pressed
                Return cachedPressedVal
            Case ImageState.Focused
                Return cachedFocusedVal
            Case ImageState.Disabled
                Return cachedDisabledVal
            Case Else
                Return cachedNormalVal
        End Select
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g = e.Graphics
        g.Clear(BackColor)

        Dim bmp As Bitmap = GetCachedBitmapForState(If(Enabled, imgState, ImageState.Disabled))
        If bmp IsNot Nothing Then
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(bmp, New Rectangle(0, 0, Width, Height), 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel)
        Else
            ' Fall back to drawing text and focus rectangle instead of image
            TextRenderer.DrawText(g, Text, Font, ClientRectangle, ForeColor, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            If Focused Then
                Dim fr = New Rectangle(4, 4, Width - 8, Height - 8)
                ControlPaint.DrawFocusRectangle(g, fr)
            End If
        End If
    End Sub
End Class
