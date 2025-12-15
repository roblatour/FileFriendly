Imports System.Diagnostics
Imports System.Windows.Documents
Imports System.Windows.Navigation

Partial Public Class LicenseWindow

    Private Sub LicenseWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Try

            Me.lbVersion.Content = "   FileFriendly v" & System.Windows.Forms.Application.ProductVersion
            While Me.lbVersion.Content.EndsWith(".0")
                Me.lbVersion.Content = Me.lbVersion.Content.Remove(Me.lbVersion.Content.Length - 2)
            End While
            Me.lbVersion.Content &= "   "

            ' ----- ABOUT TEXT (RTF) -----
            Dim aboutRtf As String = My.Resources.About

            Dim aboutRange As New TextRange(rtbAbout.Document.ContentStart, rtbAbout.Document.ContentEnd)
            Using ms As New System.IO.MemoryStream(System.Text.Encoding.Default.GetBytes(aboutRtf))
                aboutRange.Load(ms, DataFormats.Rtf)
            End Using

            WireUpHyperlinks(rtbAbout)

            ' ----- LICENSE TEXT (RTF) -----
            Dim licenseRtf As String

            licenseRtf = My.Resources.MITLicense

            Dim licenseRange As New TextRange(rtbLicense.Document.ContentStart, rtbLicense.Document.ContentEnd)
            Using ms As New System.IO.MemoryStream(System.Text.Encoding.Default.GetBytes(licenseRtf))
                licenseRange.Load(ms, DataFormats.Rtf)
            End Using

            ' If license RTF can also contain links, wire them too:
            WireUpHyperlinks(rtbLicense)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub WireUpHyperlinks(ByVal rtb As RichTextBox)

        If rtb Is Nothing OrElse rtb.Document Is Nothing Then
            Exit Sub
        End If

        Dim start As TextPointer = rtb.Document.ContentStart
        Dim [end] As TextPointer = rtb.Document.ContentEnd

        Dim position As TextPointer = start
        While position IsNot Nothing AndAlso position.CompareTo([end]) < 0
            If position.GetPointerContext(LogicalDirection.Forward) = TextPointerContext.ElementStart Then
                Dim element As TextElement = TryCast(position.GetAdjacentElement(LogicalDirection.Forward), TextElement)
                Dim link As Hyperlink = TryCast(element, Hyperlink)
                If link IsNot Nothing Then
                    ' Ensure we only hook once
                    RemoveHandler link.RequestNavigate, AddressOf Hyperlink_RequestNavigate
                    AddHandler link.RequestNavigate, AddressOf Hyperlink_RequestNavigate

                    ' Some RTF-to-FlowDocument importers may not set NavigateUri; fall back to Inlines text
                    RemoveHandler link.Click, AddressOf Hyperlink_Click
                    AddHandler link.Click, AddressOf Hyperlink_Click
                End If
            End If
            position = position.GetNextContextPosition(LogicalDirection.Forward)
        End While

    End Sub

    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)

        Try
            Dim target As String = e.Uri.AbsoluteUri
            If String.IsNullOrWhiteSpace(target) Then
                Exit Sub
            End If

            Dim psi As New ProcessStartInfo With {
                .FileName = target,
                .UseShellExecute = True
            }
            Process.Start(psi)
        Catch
            ' Swallow or optionally log
        End Try

        e.Handled = True

    End Sub

    Private Sub Hyperlink_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)

        ' Fallback if RequestNavigate is not raised / NavigateUri is missing
        Dim link As Hyperlink = TryCast(sender, Hyperlink)
        If link Is Nothing Then
            Exit Sub
        End If

        Dim target As String

        If link.NavigateUri IsNot Nothing Then
            target = link.NavigateUri.ToString()
        Else
            Dim text As String = New TextRange(link.ContentStart, link.ContentEnd).Text
            target = text.Trim()
        End If

        If String.IsNullOrWhiteSpace(target) Then
            Exit Sub
        End If

        Try
            Dim psi As New ProcessStartInfo With {
                .FileName = target,
                .UseShellExecute = True
            }
            Process.Start(psi)
        Catch
            ' Swallow or optionally log
        End Try

        e.Handled = True

    End Sub

    Private Sub LicenseWindow_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        Try
            DragMove()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub CloseCommand(ByVal sender As Object, ByVal e As Object) Handles imgClose.MouseLeftButtonDown
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class