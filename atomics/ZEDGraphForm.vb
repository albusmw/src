Option Explicit On
Option Strict On

'''<summary>Display a simple ZEDGraph window (form and graph on it).</summary>
'''<remarks>Depends on ./ZEDGraphUtil.vb.</remarks>
'''<todo>Add this class also to ZEDGraphUtil.</todo>
Public Class cZEDGraphForm

    Private Hoster As System.Windows.Forms.Form = Nothing
    Private Control As ZedGraph.ZedGraphControl = Nothing

    '''<summary>Prepare.</summary>
    Private Sub Init()
        If IsNothing(Hoster) = True Then Hoster = New System.Windows.Forms.Form
        If IsNothing(Control) = True Then Control = New ZedGraph.ZedGraphControl
        Hoster.Controls.Add(Control)
        Control.Dock = Windows.Forms.DockStyle.Fill
    End Sub

    Private Function GetGraphControl() As ZedGraph.ZedGraphControl
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(Control)
        Control.GraphPane.XAxis.MajorGrid.IsVisible = True
        Control.GraphPane.XAxis.MinorGrid.IsVisible = True
        Control.GraphPane.YAxis.MajorGrid.IsVisible = True
        Control.GraphPane.YAxis.MinorGrid.IsVisible = True
        Return Control
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal Data() As Long) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
            YAxis.Add(Data(Idx))
        Next Idx
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal Data() As Double) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
        Next Idx
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, Data, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal X() As UInt32, ByVal Y() As UInt32) As ZedGraph.ZedGraphControl
        Init()
        ZEDGraphUtil.PlotXvsY(Control, "Data", X, Y, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByRef Data As Dictionary(Of Integer, UInt32)) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Each Entry As Integer In Data.Keys
            XAxis.Add(Entry)
            YAxis.Add(Data(Entry))
        Next Entry
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByRef Data As Dictionary(Of Double, Double)) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Each Entry As Double In Data.Keys
            XAxis.Add(Entry)
            YAxis.Add(Data(Entry))
        Next Entry
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

End Class