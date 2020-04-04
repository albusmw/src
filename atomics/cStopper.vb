Option Explicit On
Option Strict On

'''<summary>Simple class to log different timing events.</summary>
Public Class cStopper

    Private Watch As New System.Diagnostics.Stopwatch
    Private TimeLog As New List(Of String)
    Private MessageCache As String = String.Empty       'message in case TIC is called with the log message and TOC is not

    Public Property PadMessage As Integer = 30
    Public Property PadTime As Integer = 11

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
    Public Function [Stamp](ByVal Text As String) As String
        Watch.Stop()
        Dim Message As String = Text.PadRight(PadMessage) & " : " & Watch.ElapsedMilliseconds.ValRegIndep.PadLeft(PadTime) & " ms"
        TimeLog.Add(Message)
        MessageCache = String.Empty
        Watch.Reset()
        Watch.Start()
        Return Message
    End Function

    '================================================================================

    '''<summary>Get the timing log.</summary>
    Public Function GetLog() As List(Of String)
        Return TimeLog
    End Function

End Class