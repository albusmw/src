Option Explicit On
Option Strict On

'''<summary>Display a simple ZEDGraph window (form and graph on it).</summary>
'''<remarks>Depends on ./ZEDGraphUtil.vb.</remarks>
'''<todo>Add this class also to ZEDGraphUtil.</todo>
Public Class cZEDGraphForm

    '''<summary>The form that shall be displayed.</summary>
    Public Hoster As System.Windows.Forms.Form = Nothing
    '''<summary>The ZED graph control inside the form.</summary>
    Private zgcMain As ZedGraph.ZedGraphControl = Nothing
    '''<summary>The ZED graph service (from file ZEDGraphService.vb).</summary>
    Public Plotter As cZEDGraphService = Nothing

    '''<summary>Prepare.</summary>
    Public Sub Init()
        If IsNothing(Hoster) = True Then Hoster = New System.Windows.Forms.Form
        If IsNothing(zgcMain) = True Then
            zgcMain = New ZedGraph.ZedGraphControl
            Hoster.Controls.Add(zgcMain)
            zgcMain.Dock = Windows.Forms.DockStyle.Fill
            Plotter = New cZEDGraphService(zgcMain)
        End If
    End Sub

    Private Function GetGraphControl() As ZedGraph.ZedGraphControl
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(zgcMain)
        zgcMain.GraphPane.XAxis.MajorGrid.IsVisible = True
        zgcMain.GraphPane.XAxis.MinorGrid.IsVisible = True
        zgcMain.GraphPane.YAxis.MajorGrid.IsVisible = True
        zgcMain.GraphPane.YAxis.MinorGrid.IsVisible = True
        Return zgcMain
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByVal Data() As Long) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
            YAxis.Add(Data(Idx))
        Next Idx
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByVal Data() As Double) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
        Next Idx
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, Data, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data with a logarithmic Y axis.</summary>
    Public Function PlotData(ByVal PlotName As String, ByVal Data() As Double, ByVal LineColor As Drawing.Color) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
        Next Idx
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, Data, New ZEDGraphUtil.sGraphStyle(LineColor), XAxis(0), XAxis(XAxis.Count - 1))
        Dim RetVal As ZedGraph.ZedGraphControl = GetGraphControl()
        Return RetVal
    End Function

    '''<summary>Plot data with a logarithmic Y axis.</summary>
    Public Function PlotDataLog(ByVal PlotName As String, ByVal Data() As Double, ByVal LineColor As Drawing.Color) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
        Next Idx
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, Data, New ZEDGraphUtil.sGraphStyle(LineColor), XAxis(0), XAxis(XAxis.Count - 1))
        Dim RetVal As ZedGraph.ZedGraphControl = GetGraphControl()
        RetVal.GraphPane.YAxis.Type = ZedGraph.AxisType.Log
        Return RetVal
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByVal X() As UInt32, ByVal Y() As UInt32) As ZedGraph.ZedGraphControl
        Init()
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, X, Y, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByVal X() As Double, ByVal Y() As Double) As ZedGraph.ZedGraphControl
        Init()
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, X, Y, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), Double.NaN, Double.NaN)
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByRef Data As Dictionary(Of Integer, UInt32)) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Each Entry As Integer In Data.Keys
            XAxis.Add(Entry)
            YAxis.Add(Data(Entry))
        Next Entry
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

    '''<summary>Plot data.</summary>
    Public Function PlotData(ByVal PlotName As String, ByRef Data As Dictionary(Of Double, Double)) As ZedGraph.ZedGraphControl
        Init()
        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Each Entry As Double In Data.Keys
            XAxis.Add(Entry)
            YAxis.Add(Data(Entry))
        Next Entry
        ZEDGraphUtil.PlotXvsY(zgcMain, PlotName, XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), XAxis(0), XAxis(XAxis.Count - 1))
        Return GetGraphControl()
    End Function

End Class