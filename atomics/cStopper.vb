Option Explicit On
Option Strict On

'''<summary>Simple class to log different timing events.</summary>
Public Class cStopper

    Private Watch As New System.Diagnostics.Stopwatch
    Private TimeLog As New List(Of String)
    Private MessageCache As String = String.Empty       'message in case TIC is called with the log message and TOC is not

    Public Property PadMessage As Integer = 30
    Public Property PadTime As Integer = 11

    Private MyProc As Process = Process.GetCurrentProcess

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
    Public Function Toc() As Long
        Dim ElapsedMilliseconds As Long = -1
        Stamp(MessageCache, ElapsedMilliseconds)
        Return ElapsedMilliseconds
    End Function

    '''<summary>Log the timing and restart the watch again.</summary>
    '''<returns>Elapes time [ms].</returns>
    Public Function Toc(ByVal Text As String) As Long
        Dim ElapsedMilliseconds As Long = -1
        Stamp(Text, ElapsedMilliseconds)
        Return ElapsedMilliseconds
    End Function

    '''<summary>Log the timing and restart the watch again.</summary>
    '''Dim ElapsedMilliseconds As Long = -1
    Public Function [Stamp](ByVal Text As String) As String
        Dim ElapsedMilliseconds As Long = -1
        Return Stamp(Text, ElapsedMilliseconds)
    End Function

    '''<summary>Log the timing and restart the watch again.</summary>
    Public Function [Stamp](ByVal Text As String, ByRef ElapsedMilliseconds As Long) As String
        Return Stamp(Text, False, ElapsedMilliseconds)
    End Function

    '''<summary>Log the timing and restart the watch again.</summary>
    '''<param name="Text">Text to add to stamp.</param>
    '''<param name="LogMemory">Log memory usage (SLOW!!!!!).</param>
    Public Function [Stamp](ByVal Text As String, ByVal LogMemory As Boolean) As String
        Dim ElapsedMilliseconds As Long = -1
        Return Stamp(Text, LogMemory, ElapsedMilliseconds)
    End Function

    '''<summary>Log the timing and restart the watch again.</summary>
    '''<param name="Text">Text to add to stamp.</param>
    '''<param name="LogMemory">Log memory usage (SLOW!!!!!).</param>
    Public Function [Stamp](ByVal Text As String, ByVal LogMemory As Boolean, ByRef ElapsedMilliseconds As Long) As String
        Watch.Stop()
        ElapsedMilliseconds = Watch.ElapsedMilliseconds
        Dim Message As String = Text.PadRight(PadMessage) & " : " & ElapsedMilliseconds.ValRegIndep.PadLeft(PadTime) & " ms"
        If LogMemory = True Then Message &= "- memory: " & Format(Process.GetCurrentProcess.PrivateMemorySize64 / 1048576, "0.0") & " MByte"
        TimeLog.Add(Format(Now, "HH.mm.ss:fff") & "|" & Message)
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