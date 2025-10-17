Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class RoundedProgressBar
    Inherits Control

    Private progMin As Integer = 0
    Private progMax As Integer = 100
    Private progValue As Integer = 0
    Private borderRadiusVal As Integer = 2
    Private borderWidthVal As Integer = 1
    Private borderColorVal As Color = Color.LightGray
    Private barColorVal As Color = SystemColors.Highlight
    Private barBackColorVal As Color = Color.Transparent

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        DoubleBuffered = True
        Size = New Size(200, 24)
    End Sub

    <DefaultValue(0)> Public Property Minimum As Integer
        Get
            Return progMin
        End Get
        Set(value As Integer)
            progMin = value
            If progMax < progMin Then progMax = progMin
            If progValue < progMin Then progValue = progMin
            Invalidate()
        End Set
    End Property

    <DefaultValue(100)> Public Property Maximum As Integer
        Get
            Return progMax
        End Get
        Set(value As Integer)
            progMax = Math.Max(value, progMin)
            If progValue > progMax Then progValue = progMax
            Invalidate()
        End Set
    End Property

    <DefaultValue(0)> Public Property Value As Integer
        Get
            Return progValue
        End Get
        Set(v As Integer)
            progValue = Math.Min(Math.Max(v, progMin), progMax)
            Invalidate()
        End Set
    End Property

    <DefaultValue(10)> Public Property Radius As Integer
        Get
            Return borderRadiusVal
        End Get
        Set(v As Integer)
            borderRadiusVal = Math.Max(0, v)
            Invalidate()
        End Set
    End Property

    <DefaultValue(2)> Public Property BorderWidth As Integer
        Get
            Return borderWidthVal
        End Get
        Set(v As Integer)
            borderWidthVal = Math.Max(0, v)
            Invalidate()
        End Set
    End Property

    Public Property BorderColor As Color
        Get
            Return borderColorVal
        End Get
        Set(v As Color)
            borderColorVal = v
            Invalidate()
        End Set
    End Property

    Public Property BarColor As Color
        Get
            Return barColorVal
        End Get
        Set(v As Color)
            barColorVal = v
            Invalidate()
        End Set
    End Property

    Public Property BarBackColor As Color
        Get
            Return barBackColorVal
        End Get
        Set(v As Color)
            barBackColorVal = v
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g = e.Graphics

        g.SmoothingMode = SmoothingMode.AntiAlias

        ' Shift offset by 0.5px for crisp strokes as otherwise strokes are blurry (similar to Rainmeter)
        Dim half As Single = 0.5F
        Dim outer = New RectangleF(0 + half, 0 + half, Width - 1, Height - 1)

        ' Inset inner by half the border width so the fill abuts the inner visible edge
        Dim inset = borderWidthVal / 2.0F
        Dim inner = New RectangleF(outer.X + inset, outer.Y + inset, Math.Max(0, outer.Width - inset * 2), Math.Max(0, outer.Height - inset * 2))
        Dim innerRadius = Math.Max(0, borderRadiusVal - borderWidthVal / 1.5)

        Using trackPath = GetRoundedRectPath(Rectangle.Round(inner), innerRadius)
            Using brushBack As New SolidBrush(barBackColorVal)
                g.FillPath(brushBack, trackPath)
            End Using
        End Using

        Dim range = Math.Max(1, progMax - progMin)
        Dim percent = (progValue - progMin) / CSng(range)
        Dim fillWidth = CInt(Math.Round(inner.Width * percent))

        If fillWidth > 0 Then
            ' Clip to inner rounded path
            Using clipPath = GetRoundedRectPath(Rectangle.Round(inner), CInt(Math.Max(0, innerRadius)))
                g.SetClip(clipPath)
                g.SmoothingMode = SmoothingMode.None

                Dim innerXR = inner.X
                Dim innerYR = inner.Y
                Dim innerHR = inner.Height

                Dim fillW = Math.Min(fillWidth + 1, inner.Width)
                Dim fillH = innerHR

                Dim fillRect As New Rectangle(innerXR, innerYR, fillW, fillH)

                Using brushFill As New SolidBrush(barColorVal)
                    g.FillRectangle(brushFill, fillRect)
                End Using

                g.SmoothingMode = SmoothingMode.AntiAlias
                g.ResetClip()
            End Using
        End If

        ' Draw rounded border using a centered stroke positioned by sub-pixel offset
        If borderWidthVal > 0 Then
            Using pen As New Pen(borderColorVal, borderWidthVal)
                pen.Alignment = PenAlignment.Center
                Using borderPath = GetRoundedRectPath(Rectangle.Round(outer), borderRadiusVal)
                    g.DrawPath(pen, borderPath)
                End Using
            End Using
        End If
    End Sub

    Private Function GetRoundedRectPath(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        If radius <= 0 Then
            path.AddRectangle(rect)
            path.CloseFigure()
            Return path
        End If

        Dim d = radius * 2
        path.AddArc(rect.X, rect.Y, d, d, 180, 90)
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90)
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90)
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Invalidate()
    End Sub
End Class
