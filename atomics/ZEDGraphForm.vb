Option Explicit On
Option Strict On

'''<summary>Display a simple ZEDGraph window (form and graph on it).</summary>
'''<remarks>Depends on ./ZEDGraphUtil.vb.</remarks>
'''<todo>Add this class also to ZEDGraphUtil.</todo>
Public Class ZEDGraphForm

    '''<summary>Plot data.</summary>
    Public Shared Function PlotData(ByVal Data() As Long) As ZedGraph.ZedGraphControl

        Dim Hoster As New System.Windows.Forms.Form
        Dim Control As New ZedGraph.ZedGraphControl

        Hoster.Controls.Add(Control)
        Control.Dock = Windows.Forms.DockStyle.Fill

        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
            YAxis.Add(Data(Idx))
        Next Idx
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(Control)

        Return Control

    End Function

    '''<summary>Plot data.</summary>
    Public Shared Function PlotData(ByVal Data() As Double) As ZedGraph.ZedGraphControl

        Dim Hoster As New System.Windows.Forms.Form
        Dim Control As New ZedGraph.ZedGraphControl

        Hoster.Controls.Add(Control)
        Control.Dock = Windows.Forms.DockStyle.Fill

        Dim XAxis As New List(Of Double)
        For Idx As Integer = 0 To Data.GetUpperBound(0)
            XAxis.Add(Idx)
        Next Idx
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, Data, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red), XAxis(0), XAxis(XAxis.Count - 1))
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(Control)

        Return Control

    End Function

    '''<summary>Plot data.</summary>
    Public Shared Function PlotData(ByVal X() As UInt32, ByVal Y() As UInt32) As ZedGraph.ZedGraphControl

        Dim Hoster As New System.Windows.Forms.Form
        Dim Control As New ZedGraph.ZedGraphControl

        Hoster.Controls.Add(Control)
        Control.Dock = Windows.Forms.DockStyle.Fill

        ZEDGraphUtil.PlotXvsY(Control, "Data", X, Y, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots))
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(Control)

        Return Control

    End Function

    '''<summary>Plot data.</summary>
    Public Shared Function PlotData(ByRef Data As Dictionary(Of Integer, UInt32)) As ZedGraph.ZedGraphControl

        Dim Hoster As New System.Windows.Forms.Form
        Dim Control As New ZedGraph.ZedGraphControl

        Hoster.Controls.Add(Control)
        Control.Dock = Windows.Forms.DockStyle.Fill

        Dim XAxis As New List(Of Double)
        Dim YAxis As New List(Of Double)
        For Each Entry As Integer In Data.Keys
            XAxis.Add(Entry)
            YAxis.Add(Data(Entry))
        Next Entry
        ZEDGraphUtil.PlotXvsY(Control, "Data", XAxis.ToArray, YAxis.ToArray, New ZEDGraphUtil.sGraphStyle(Drawing.Color.Red, ZEDGraphUtil.sGraphStyle.eCurveMode.Dots), XAxis(0), XAxis(XAxis.Count - 1))
        Hoster.Show()
        ZEDGraphUtil.MaximizePlotArea(Control)

        Return Control

    End Function

End Class