Option Explicit On
Option Strict On

'''<summary>Simple class to log different timing events.</summary>
Public Class cStopper

    Private Watch As New Stopwatch
    Private TimeLog As New List(Of String)
    Private MessageCache As String = String.Empty       'message in case TIC is called with the log message and TOC is not

    '''<summary>Constructor.</summary>
    Public Sub New()
        Watch.Reset() : Watch.Start()
    End Sub

    '================================================================================

    '''<summary>(Re-)start the stopwatch.</summary>
    Public Sub Tic()
        [Start]()
    End Sub

    '''<summary>(Re-)start the stopwatch.</summary>
    Public Sub Tic(ByVal Text As String)
        [Start]()
        MessageCache = Text
    End Sub

    '''<summary>(Re-)start the stopwatch.</summary>
    Public Sub [Start]()
        Watch.Reset() : Watch.Start()
    End Sub


    '================================================================================

    '''<summary>Log the timing and restart the watch again.</summary>
    Public Sub Toc()
        Stamp(MessageCache)
    End Sub

    '''<summary>Log the timing and restart the watch again.</summary>
    Public Sub Toc(ByVal Text As String)
        Stamp(Text)
    End Sub

    '''<summary>Log the timing and restart the watch again.</summary>
    Public Sub [Stamp](ByVal Text As String)
        Watch.Stop()
        TimeLog.Add(Text.PadRight(30) & " : " & Watch.ElapsedMilliseconds.ToString.Trim.PadLeft(11) & " ms")
        MessageCache = String.Empty
        Watch.Reset()
        Watch.Start()
    End Sub

    '================================================================================

    '''<summary>Get the timing log.</summary>
    Public Function GetLog() As String()
        Return TimeLog.ToArray
    End Function

End Class