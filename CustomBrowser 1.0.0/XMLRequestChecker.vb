Imports Gecko.Observers
Imports Gecko.Net
Imports System.Text

Public Class XMLRequestChecker
    Inherits BaseHttpModifyRequestObserver
    Friend Event CaptureResponseData(ByVal captured As String)

    Protected Overrides Sub ObserveRequest(channel As HttpChannel)
        Dim controlString As String
        Select Case BrowserID
            Case 1
                controlString = "suggest-geo"
            Case 2
                controlString = "autosuggest"
        End Select
        If channel IsNot Nothing Then
            If channel.Uri.AbsolutePath.Contains(controlString) Then
                Dim TC As TraceableChannel = channel.CastToTraceableChannel
                Dim Stream As StreamListenerTee = New StreamListenerTee
                AddHandler Stream.Completed, AddressOf StreamCompleted
                TC.SetNewListener(Stream)
            End If
        End If
    End Sub

    Private Sub StreamCompleted(ByVal sender As Object, ByVal e As EventArgs)
        If TypeOf sender Is StreamListenerTee Then
            Dim Stream As StreamListenerTee = TryCast(sender, StreamListenerTee)
            Dim data As Byte() = Stream.GetCapturedData
            Dim StringData As String = Encoding.UTF8.GetString(data)
            RaiseEvent CaptureResponseData(StringData)
        End If
    End Sub
End Class
