Option Explicit On
Option Strict On

'''<summary>Logging stack class.</summary>
Public Class cLogging

    '''<summary>Level of logging.</summary>
    Public Enum eLogLevel
        [Debug]
        [Info]
        [Timing]
        [Warning]
        [Error]
    End Enum

    Public Shared LogLevelText As String() = {"Debug  ", "Info   ", "Timing ", "Warning", "Error  "}

    '''<summary>One logging entry.</summary>
    Public Structure sLogEntry
        Public Moment As DateTime
        Public Level As eLogLevel
        Public Message As String
        Public Sub New(ByVal NewMoment As DateTime, ByVal NewLevel As eLogLevel, ByVal NewMessage As String)
            Moment = NewMoment
            Level = NewLevel
            Message = NewMessage
        End Sub
        Public Function Verbose() As String
            Return Format(Now, "HH:mm:ss") & "|" & LogLevelText(Level) & "|" & Message
        End Function
    End Structure

    '''<summary>One tic-toc entry.</summary>
    Public Structure sTicToc
        Public Stopper As Stopwatch
        Public Context As String
        Public Sub New(ByVal NewContext As String)
            Stopper = New Stopwatch
            Context = NewContext
            Stopper.Reset() : Stopper.Start()
        End Sub
    End Structure

    '''<summary>Dictionary with all log entries.</summary>
    Private LogList As New Dictionary(Of Long, sLogEntry)
    Private LogListPtr As Long = -1

    '''<summary>Dictionary with all running stopwatches.</summary>
    Private StopList As New Dictionary(Of Long, sTicToc)
    Private StopListPtr As Long = -1

    '''<summary>Add the entry to the log.</summary>
    Public Function Log(ByVal Message As String) As Long
        Return Log(New sLogEntry(Now, eLogLevel.Info, Message))
    End Function

    '''<summary>Add the entry to the log.</summary>
    Public Function LogError(ByVal Message As String) As Long
        Return Log(New sLogEntry(Now, eLogLevel.Error, Message))
    End Function

    '''<summary>Most-inner log function.</summary>
    Public Function Log(ByVal LogEntry As sLogEntry) As Long
        LogListPtr += 1
        LogList.Add(LogListPtr, LogEntry)
        Return LogListPtr
    End Function

    '''<summary>Get all new content since the last call.</summary>
    '''<param name="LastListPtr">Last read entry pointer - will be changed to the new one.</param>
    Public Function GetNewContent(ByRef LastListPtr As Long) As List(Of sLogEntry)
        If LogListPtr > LastListPtr Then
            Dim RetList As New List(Of sLogEntry)
            Dim PtrToStopAt As Long = LogListPtr
            For Ptr As Long = LastListPtr + 1 To PtrToStopAt
                RetList.Add(LogList(Ptr))
            Next Ptr
            LastListPtr = PtrToStopAt
            Return RetList
        Else
            'Nothing to return
            Return New List(Of sLogEntry)({})
        End If
    End Function

    '''<summary>Call when entering a certain function.</summary>
    '''<returns>Dictionary pointer to running stopwatch.</returns>
    Public Function Tic(ByVal Context As String) As Long
        Log(New sLogEntry(Now, eLogLevel.Timing, "Tic for <" & Context & ">"))
        StopListPtr += 1
        StopList.Add(StopListPtr, New sTicToc(Context))
        Return StopListPtr
    End Function

    '''<summary>Call when a certain action ends.</summary>
    Public Function Toc(ByVal StopperPtr As Long) As Long
        If StopList.ContainsKey(StopperPtr) = False Then Return -1
        StopList(StopperPtr).Stopper.Stop()
        Dim ElapsedMilliseconds As Long = StopList(StopperPtr).Stopper.ElapsedMilliseconds
        StopList(StopperPtr).Stopper.Reset()
        Log(New sLogEntry(Now, eLogLevel.Timing, "Toc for <" & StopList(StopperPtr).Context & ">: " & ElapsedMilliseconds.ToString.Trim & " ms"))
        Return ElapsedMilliseconds
    End Function

End Class

'''<summary>DO NOT USE ....</summary>
<Obsolete("Do not use", True)>
Public Class cTicTocLogging

    Dim Content As New List(Of String)

    '''<summary>Add a free text to the log.</summary>
    Public Sub Add(ByVal Text As String)
        Content.Add(Format(Now, "hh:mm:ss.fff" & "|") & Text)
    End Sub

    Public Sub ShowLog()
        ShowLog(12)
    End Sub

    Public Sub ShowLog(ByVal FontSize As Integer)
        Dim FormToShow As New Windows.Forms.Form
        Dim TextB As New Windows.Forms.TextBox
        With TextB
            .Multiline = True
            .Dock = Windows.Forms.DockStyle.Fill
            .ScrollBars = Windows.Forms.ScrollBars.Both
            .Text = Join(Content.ToArray, System.Environment.NewLine)
            .Font = New Drawing.Font("Courier New", FontSize)
        End With
        FormToShow.Controls.Add(TextB)
        FormToShow.WindowState = Windows.Forms.FormWindowState.Maximized
        FormToShow.Show()
    End Sub

End Class
