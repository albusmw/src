Option Explicit On
Option Strict On

'''<summary>This class procides a simple Tic-Toc logging and a small form to show the log content.</summary>
Public Class cLogging

    Private MaxCallDepth As Integer = 0
    Private MyLastEntry As String = String.Empty
    Private Content As New List(Of String)
    Private Stopper(MaxCallDepth) As Stopwatch
    Private StopperPtr As Integer = 0
    Private StopperContext As String = String.Empty

    '''<summary>Call when entering a certain function.</summary>
    Public Sub Tic(ByVal Context As String)
        Add("Starting <" & Context & ">")
        StopperContext = Context
        StopperPtr += 1
        Stopper(StopperPtr).Reset()
        Stopper(StopperPtr).Start()
    End Sub

    Public ReadOnly Property LastEntry() As String
        Get
            Return MyLastEntry
        End Get
    End Property

    '''<summary>Call when a certain action ends.</summary>
    Public Sub Toc()
        Stopper(StopperPtr).Stop()
        Add("Elapsed time for <" & StopperContext & ">: " & Stopper(StopperPtr).ElapsedMilliseconds.ToString.Trim)
        StopperPtr -= 1
    End Sub

    '''<summary>Add a free text to the log.</summary>
    Public Sub Add(ByVal Text As String)
        Content.Add(Format(Now, "hh:mm:ss.fff" & "|") & Text)
        MyLastEntry = Text
    End Sub

    Public Sub ShowLog()
        Dim FormToShow As New Windows.Forms.Form
        Dim TextB As New Windows.Forms.TextBox
        With TextB
            .Multiline = True
            .Dock = Windows.Forms.DockStyle.Fill
            .ScrollBars = Windows.Forms.ScrollBars.Both
            .Text = Join(Content.ToArray, System.Environment.NewLine)
            .Font = New Drawing.Font("Courier New", 16)
        End With
        FormToShow.Controls.Add(TextB)
        FormToShow.WindowState = Windows.Forms.FormWindowState.Maximized
        FormToShow.Show()
    End Sub

    Public Sub New()
        For Idx As Integer = 0 To Stopper.GetUpperBound(0)
            Stopper(Idx) = New Stopwatch : Stopper(Idx).Reset() : Stopper(Idx).Stop()
        Next Idx
    End Sub

End Class
