Option Explicit On
Option Strict On

'Move content to C:\GIT\src\atomics\ZedGraphService.vb
Public Class ZEDGraphUtil

    '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
    '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
    '''<param name="X">Vector of X axis values.</param>
    '''<param name="Y">Vector of Y axis values.</param>
    '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
    Public Shared Sub PlotXvsY(ByRef Graph As ZedGraph.ZedGraphControl, ByRef CurveName As String, ByRef X() As Double, ByRef Y() As Double, ByVal Style As sGraphStyle, ByVal X_Min As Double, ByVal X_Max As Double)

        'If the curve list is empty, add the new surve
        Dim EditedLine As ZedGraph.LineItem = Nothing
        If Graph.GraphPane.CurveList.Count = 0 Then
            EditedLine = Graph.GraphPane.AddCurve(CurveName, X, Y, Style.LineColor, Style.DotStyle)
        Else
            For Each Item As ZedGraph.LineItem In Graph.GraphPane.CurveList
                If Item.Label.Text = CurveName Then
                    EditedLine = Item
                    EditedLine.Clear()
                    'Protect against corrupt X and Y data
                    If (IsNothing(X) = False) And (IsNothing(Y) = False) Then
                        If X.Length = Y.Length Then
                            For Idx As Integer = 0 To X.GetUpperBound(0)
                                Item.AddPoint(X(Idx), Y(Idx))
                            Next
                        End If
                    End If
                End If
            Next Item
        End If

        'Set curve style
        If IsNothing(EditedLine) = False Then Style.ConfigureLineItem(EditedLine)

        With Graph
            If Double.IsNaN(X_Min) = False Then
                .GraphPane.XAxis.Scale.MinAuto = False
                .GraphPane.XAxis.Scale.Min = X_Min
            Else
                .GraphPane.XAxis.Scale.MinAuto = True
            End If
            If Double.IsNaN(X_Max) = False Then
                .GraphPane.XAxis.Scale.MaxAuto = False
                .GraphPane.XAxis.Scale.Max = X_Max
            Else
                .GraphPane.XAxis.Scale.MaxAuto = True
            End If
            .GraphPane.YAxis.Scale.MaxAuto = True
            .GraphPane.YAxis.Scale.MinAuto = True
            .AxisChange()
            .Invalidate()
            .Refresh()
        End With

    End Sub

    '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
    '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
    '''<param name="X">Vector of X axis values.</param>
    '''<param name="Y">Vector of Y axis values.</param>
    '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
    Public Shared Sub PlotXvsY(ByRef Graph As ZedGraph.ZedGraphControl, ByRef CurveName As String, ByRef X() As UInt32, ByRef Y() As UInt32, ByVal Style As sGraphStyle)

        'If the curve list is empty, add the new surve
        Dim EditedLine As ZedGraph.LineItem = Nothing
        If Graph.GraphPane.CurveList.Count = 0 Then
            EditedLine = Graph.GraphPane.AddCurve(CurveName, Array.ConvertAll(X, New Converter(Of UInt32, Double)(AddressOf DoubleToUInt32)), Array.ConvertAll(Y, New Converter(Of UInt32, Double)(AddressOf DoubleToUInt32)), Style.LineColor, Style.DotStyle)
        Else
            For Each Item As ZedGraph.LineItem In Graph.GraphPane.CurveList
                If Item.Label.Text = CurveName Then
                    EditedLine = Item
                    EditedLine.Clear()
                    'Protect against corrupt X and Y data
                    If (IsNothing(X) = False) And (IsNothing(Y) = False) Then
                        If X.Length = Y.Length Then
                            For Idx As Integer = 0 To X.GetUpperBound(0)
                                Item.AddPoint(X(Idx), Y(Idx))
                            Next
                        End If
                    End If
                End If
            Next Item
        End If

        'Set curve style
        If IsNothing(EditedLine) = False Then Style.ConfigureLineItem(EditedLine)

        With Graph
            .GraphPane.XAxis.Scale.MinAuto = True
            .GraphPane.XAxis.Scale.MaxAuto = True
            .GraphPane.YAxis.Scale.MaxAuto = True
            .GraphPane.YAxis.Scale.MinAuto = True
            .AxisChange()
            .Invalidate()
            .Refresh()
        End With

    End Sub

    Public Shared Sub SetCaptions(ByRef Graph As ZedGraph.ZedGraphControl, ByVal Title As String, ByVal XAxis As String, ByVal YAxis As String)
        Graph.GraphPane.Title.Text = Title
        Graph.GraphPane.XAxis.Title.Text = XAxis
        Graph.GraphPane.YAxis.Title.Text = YAxis
    End Sub

    Public Shared Sub MaximizePlotArea(ByRef Graph As ZedGraph.ZedGraphControl)
        With Graph
            .GraphPane.Margin.All = 0
            .GraphPane.XAxis.Title.FontSpec.Size = 12
            .GraphPane.YAxis.Title.FontSpec.Size = 12
            .GraphPane.TitleGap = 0
            .GraphPane.Title.FontSpec.Size = 14
            .GraphPane.Legend.FontSpec.Size = 10
        End With
    End Sub

    '''<summary>Structure to describe the line and point style.</summary>
    Public Structure sGraphStyle

        '''<summary>Display lines and/or dots.</summary>
        Public Enum eCurveMode
            '''<summary>Line only.</summary>
            Lines
            '''<summary>Dots only.</summary>
            Dots
            '''<summary>Lines and dots.</summary>
            LinesAndPoints
        End Enum

        '''<summary>Style of the curve.</summary>
        Public CurveMode As eCurveMode
    '''<summary>Line color to select.</summary>
    Public LineColor As Drawing.Color
    '''<summary>Line width.</summary>
    Public LineWidth As Single
    '''<summary>Line color to select.</summary>
    Public LineStyle As Drawing.Drawing2D.DashStyle
    '''<summary>Color of the dots = samples.</summary>
    Public DotColor As Drawing.Color
    '''<summary>Style of the dots.</summary>
    Public DotStyle As ZedGraph.SymbolType

    '''<summary>Create the default style with a defined color.</summary>
    '''<param name="Color">Color to apply.</param>
    Public Sub New(ByVal Color As Drawing.Color, Optional ByVal LineWidth As Integer = 1)
      Init(Color)
      Me.LineWidth = LineWidth
    End Sub

    '''<summary>Determine if the plot line is visible.</summary>
    Public ReadOnly Property LineVisible() As Boolean
      Get
        Select Case CurveMode
          Case eCurveMode.Lines, eCurveMode.LinesAndPoints
            Return True
          Case Else
            Return False
        End Select
      End Get
    End Property

    '''<summary>Determine if the plot dots are visible.</summary>
    Public ReadOnly Property DotVisible() As Boolean
      Get
        Select Case CurveMode
          Case eCurveMode.Dots, eCurveMode.LinesAndPoints
            Return True
          Case Else
            Return False
        End Select
      End Get
    End Property

    '''<summary>Create the default style with a defined color and switch between line or dot plot.</summary>
    '''<param name="NewColor">Color to apply.</param>
    '''<param name="NewCurveMode">CurveMode (plot line and/or dots).</param>
    Public Sub New(ByVal NewColor As Drawing.Color, ByVal NewCurveMode As eCurveMode, Optional ByVal NewLineWidth As Integer = 1)
      Init(NewColor)
      CurveMode = NewCurveMode
      LineWidth = NewLineWidth
    End Sub

    '''<summary>Create the default style with a defined line and point color.</summary>
    '''<param name="NewLineColor">Color of the lines.</param>
    '''<param name="NewDotColor">Color of the points.</param>
    Public Sub New(ByVal NewLineColor As Drawing.Color, ByVal NewDotColor As Drawing.Color, Optional ByVal NewLineWidth As Integer = 1)
      Init(NewLineColor)
      DotColor = NewDotColor
      LineWidth = NewLineWidth
    End Sub

    Private Sub Init(ByRef Color As Drawing.Color)
      With Me
        .CurveMode = eCurveMode.Lines
        .LineColor = Color
        .LineStyle = Drawing.Drawing2D.DashStyle.Solid
        .DotColor = Color
        .DotStyle = ZedGraph.SymbolType.Circle
      End With
    End Sub

    '''<summary>Apply the configuration of this object to the given line item.</summary>
    Public Sub ConfigureLineItem(ByRef LineItem As ZedGraph.LineItem)

      'Configure the line
      With LineItem.Line
        .Color = LineColor
        If LineVisible = True Then .Width = LineWidth
        .IsVisible = LineVisible
        .Style = LineStyle
      End With

      'Configure the symbol
      With LineItem.Symbol
        .Size = LineWidth + 1
        .IsVisible = DotVisible
        .Border.Color = DotColor
        .Fill.Color = DotColor
        .Fill.IsVisible = DotVisible
        .Fill.Type = ZedGraph.FillType.Solid
      End With

    End Sub

  End Structure

  '''<summary>Delete all displayed graphs.</summary>
  '''<remarks>IsNothing was added to protect against problems when calling with un-associated graph control.</remarks>
  Public Sub Clear()
    If IsNothing(CurveList) = False Then CurveList.Clear() 'clear curve list
    If IsNothing(MainGraph.GraphPane.CurveList) = False Then MainGraph.GraphPane.CurveList.Clear() 'clear all displayed curves
  End Sub

  '''<summary>All curves added with the plot functions.</summary>
  Private CurveList As New Dictionary(Of String, ZedGraph.LineItem)

  '''<summary>Graph the library is currently attached to.</summary>
  Private MainGraph As ZedGraph.ZedGraphControl

  '''<summary>Establish the link between the graph component and this class.</summary>
  '''<param name="Graph">Graph to attache class to.</param>
  Public Sub AttachToGraph(ByRef Graph As ZedGraph.ZedGraphControl)
    MainGraph = Graph
    MainGraph.GraphPane.XAxis.Title.Text = "Date"
    MainGraph.GraphPane.XAxis.Type = ZedGraph.AxisType.Date
    MainGraph.GraphPane.XAxis.Scale.Format = "dd.MM.yyyy"
    MainGraph.GraphPane.XAxis.Scale.MajorUnit = ZedGraph.DateUnit.Day
    MainGraph.GraphPane.XAxis.Scale.MajorStep = 7
    MainGraph.GraphPane.XAxis.Scale.MinorStep = 1
  End Sub

  '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
  '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
  '''<param name="X">Vector of X axis values.</param>
  '''<param name="Y">Vector of Y axis values.</param>
  '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
  '''<param name="Use2ndYAxis">Use the 2nd Y axis for plotting?</param>
  Public Sub PlotXvsT(ByRef CurveName As String, ByRef X() As Date, ByRef Y() As Double, ByVal Style As sGraphStyle, ByVal Use2ndYAxis As Boolean)

    If CurveList.ContainsKey(CurveName) = False Then
      CurveList.Add(CurveName, MainGraph.GraphPane.AddCurve(CurveName, Nothing, Style.LineColor, Style.DotStyle))
    End If
    UpdateXvsY(CurveName, X, Y)

    With CurveList(CurveName)

      .Tag = CurveName

      'Asign to axis
      .IsY2Axis = Use2ndYAxis

      Style.ConfigureLineItem(CurveList(CurveName))

    End With

    If X(0) > X(X.GetUpperBound(0)) Then
      MainGraph.GraphPane.XAxis.Scale.Min = New ZedGraph.XDate(X(X.GetUpperBound(0)))
      MainGraph.GraphPane.XAxis.Scale.Max = New ZedGraph.XDate(X(0))
    Else
      MainGraph.GraphPane.XAxis.Scale.Min = New ZedGraph.XDate(X(0))
      MainGraph.GraphPane.XAxis.Scale.Max = New ZedGraph.XDate(X(X.GetUpperBound(0)))
    End If
    AutoScaleYAxis()
    MainGraph.Invalidate()
    MainGraph.Refresh()

  End Sub

  Private Sub UpdateXvsY(ByRef CurveName As String, ByRef X As Date(), ByRef Y As Double())

    'If the curve name is not known, don't update
    If CurveList.ContainsKey(CurveName) = False Then Exit Sub

    'Process
    With CurveList(CurveName)
      .Clear()
      'Protect against corrupt X and Y data
      If (IsNothing(X) = False) And (IsNothing(Y) = False) Then
        If X.Length = Y.Length Then
          For Idx As Integer = 0 To X.GetUpperBound(0)
            .AddPoint(New ZedGraph.XDate(X(Idx)), Y(Idx))
          Next
        End If
      End If

    End With
  End Sub

    '''<summary>Set Y axis to auto-scale and update scaling.</summary>
    Public Sub AutoScaleYAxis()
        MainGraph.GraphPane.YAxis.Scale.MaxAuto = True
        MainGraph.GraphPane.YAxis.Scale.MinAuto = True
        MainGraph.AxisChange()
    End Sub

    Private Shared Function DoubleToUInt32(ByVal DataIn As UInt32) As Double
        Return DataIn
    End Function

End Class