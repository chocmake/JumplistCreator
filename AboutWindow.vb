Partial Public Class AboutWindow
    Inherits System.Windows.Forms.Form

    Public Sub New()
        InitializeComponent()
        ThemeLoader.TryApplyTheme(Me)

        AddHandler Me.WebsiteLinkLabel.LinkClicked, AddressOf WebsiteLinkLabelClicked
    End Sub

    Private Sub WebsiteLinkLabelClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        e.Link.Visited = True
        Dim url As String = TryCast(e.Link.LinkData, String)
        If Not String.IsNullOrEmpty(url) Then
            Try
                System.Diagnostics.Process.Start(url)
            Catch ex As Exception
            End Try
        End If
    End Sub
End Class
