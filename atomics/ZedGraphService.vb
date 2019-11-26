Option Explicit On
Option Strict On

'''<summary>Class to simplify the interop between the application and the ZedGraph graphical display.</summary>
Public Class cZEDGraphService

  '''<summary>Display lines and/or dots.</summary>
  Public Enum eCurveMode
    '''<summary>Line only.</summary>
    Lines
    '''<summary>Dots only.</summary>
    Dots
    '''<summary>Lines and dots.</summary>
    LinesAndPoints
  End Enum

  Public Enum eLineStyle
    Solid
  End Enum

  '''<summary>Graph the library is currently attached to.</summary>
  Private MainGraph As ZedGraph.ZedGraphControl
  '''<summary>All curves added with the plot functions.</summary>
  Private CurveList As New Dictionary(Of String, ZedGraph.LineItem)

  '''<summary>This class can be used to generate a sequence of curve styles which look all different.</summary>
  Public Class cLineStyleGenerator

    Private ColorList As List(Of Drawing.Color)
    Private ColorPtr As Integer = -1
    Private LineStyleList As List(Of Drawing.Drawing2D.DashStyle)
    Private LineStylePtr As Integer = -1

    Public Sub New()

      ColorPtr = -1
      ColorList = New List(Of Drawing.Color)

      ColorList.Add(Drawing.Color.Red)
      ColorList.Add(Drawing.Color.Green)
      ColorList.Add(Drawing.Color.Blue)
      ColorList.Add(Drawing.Color.Cyan)
      ColorList.Add(Drawing.Color.Magenta)
      ColorList.Add(Drawing.Color.Orange)

      ColorList.Add(Drawing.Color.DarkRed)
      ColorList.Add(Drawing.Color.DarkGreen)
      ColorList.Add(Drawing.Color.DarkBlue)
      ColorList.Add(Drawing.Color.DarkCyan)
      ColorList.Add(Drawing.Color.DarkMagenta)
      ColorList.Add(Drawing.Color.DarkOrange)

      LineStylePtr = 0
      LineStyleList = New List(Of Drawing.Drawing2D.DashStyle)
      LineStyleList.Add(Drawing.Drawing2D.DashStyle.Solid)
      LineStyleList.Add(Drawing.Drawing2D.DashStyle.Dot)
      LineStyleList.Add(Drawing.Drawing2D.DashStyle.Dash)

    End Sub

    Public Function GetNextStyle(ByVal CurveMode As eCurveMode) As sGraphStyle
      ColorPtr += 1
      If ColorPtr > ColorList.Count - 1 Then
        ColorPtr = 0
        LineStylePtr += 1
        If LineStylePtr > LineStyleList.Count - 1 Then
          LineStylePtr = 0
        End If
      End If
      Dim RetVal As New sGraphStyle(ColorList(ColorPtr), CurveMode)
      RetVal.LineStyle = LineStyleList(LineStylePtr)
      Return RetVal
    End Function

  End Class


  '''<summary>Structure to describe the line and point style.</summary>
  Public Structure sGraphStyle

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

  '''<summary>Establish the link between the graph component and this class.</summary>
  '''<param name="Graph">Graph to attache class to.</param>
  Public Sub New(ByRef Graph As ZedGraph.ZedGraphControl)
    AttachToGraph(Graph)
  End Sub

  '''<summary>Establish the link between the graph component and this class.</summary>
  '''<param name="Graph">Graph to attache class to.</param>
  Private Sub AttachToGraph(ByRef Graph As ZedGraph.ZedGraphControl)
    MainGraph = Graph
  End Sub

  '''<summary>Delete all displayed graphs.</summary>
  '''<remarks>IsNothing was added to protect against problems when calling with un-associated graph control.</remarks>
  Public Sub Clear()
    If IsNothing(CurveList) = False Then CurveList.Clear() 'clear curve list
    If IsNothing(MainGraph.GraphPane.CurveList) = False Then MainGraph.GraphPane.CurveList.Clear() 'clear all displayed curves
  End Sub

  Public Sub AutoScaleAxis()
    AutoScaleXAxis()
    AutoScaleYAxis()
  End Sub

  '''<summary>Set X axis to auto-scale and update scaling.</summary>
  Public Sub AutoScaleXAxis()
    MainGraph.GraphPane.XAxis.Scale.MaxAuto = True
    MainGraph.GraphPane.XAxis.Scale.MinAuto = True
    MainGraph.AxisChange()
  End Sub

    '''<summary>Set Y axis to auto-scale and update scaling.</summary>
    Public Sub AutoScaleYAxis()
        MainGraph.GraphPane.YAxis.Type = ZedGraph.AxisType.Linear
        MainGraph.GraphPane.YAxis.Scale.MaxAuto = True
        MainGraph.GraphPane.YAxis.Scale.MinAuto = True
        MainGraph.AxisChange()
    End Sub

    '''<summary>Set Y axis to auto-scale and update scaling.</summary>
    Public Sub AutoScaleYAxisLog()
        MainGraph.GraphPane.YAxis.Type = ZedGraph.AxisType.Log
        MainGraph.GraphPane.YAxis.Scale.MaxAuto = True
        MainGraph.GraphPane.YAxis.Scale.MinAuto = True
        MainGraph.AxisChange()
    End Sub

    '''<summary>Set X axis Min and Max and update scaling.</summary>
    Public Sub ManuallyScaleXAxis(ByVal MinValue As Double, ByVal MaxValue As Double)
    MainGraph.GraphPane.XAxis.Scale.Max = MaxValue
    MainGraph.GraphPane.XAxis.Scale.Min = MinValue
    MainGraph.AxisChange()
  End Sub

  '''<summary>Set Y axis Min and Max and update scaling.</summary>
  Public Sub ManuallyScaleYAxis(ByVal MinValue As Double, ByVal MaxValue As Double)
    MainGraph.GraphPane.YAxis.Scale.Max = MaxValue
    MainGraph.GraphPane.YAxis.Scale.Min = MinValue
    MainGraph.AxisChange()
  End Sub

  Public Sub UpdateDisplay(ByVal RunAutoScale As Boolean)
    If RunAutoScale = True Then
      MainGraph.GraphPane.XAxis.Scale.MaxAuto = True
      MainGraph.GraphPane.XAxis.Scale.MinAuto = True
      MainGraph.GraphPane.YAxis.Scale.MaxAuto = True
      MainGraph.GraphPane.YAxis.Scale.MinAuto = True
      MainGraph.AxisChange()
    End If
    MainGraph.Invalidate()
  End Sub

  '''<summary>Ensure that the display is updated.</summary>
  Public Sub ForceUpdate()
    MainGraph.Invalidate()
    MainGraph.Refresh()
  End Sub

  '''<summary>Force the X and Y axis to have same "size".</summary>
  Public Sub ForceToSquare(ByVal CenterZero As Boolean)
    With MainGraph.GraphPane

      If CenterZero = True Then

        'Adjust the scaling to make a square (after doing AutoScale and ensuring that all points are present)
        Dim HalfSpan As Double = Double.MinValue
        If Math.Abs(.XAxis.Scale.Max) > HalfSpan Then HalfSpan = Math.Abs(.XAxis.Scale.Max)
        If Math.Abs(.XAxis.Scale.Min) > HalfSpan Then HalfSpan = Math.Abs(.XAxis.Scale.Min)
        If Math.Abs(.YAxis.Scale.Max) > HalfSpan Then HalfSpan = Math.Abs(.YAxis.Scale.Max)
        If Math.Abs(.YAxis.Scale.Min) > HalfSpan Then HalfSpan = Math.Abs(.YAxis.Scale.Min)
        .XAxis.Scale.Max = HalfSpan : .XAxis.Scale.Min = -HalfSpan
        .YAxis.Scale.Max = HalfSpan : .YAxis.Scale.Min = -HalfSpan

        'Calculate a better scaling (no 5.76 as maximum but 6 instread, ...)
        .XAxis.Scale.MajorStep = 1
        .YAxis.Scale.MajorStep = .XAxis.Scale.MajorStep
        .XAxis.Scale.MinorStep = .XAxis.Scale.MajorStep / 5
        .YAxis.Scale.MinorStep = .XAxis.Scale.MinorStep

      Else

        Dim XSpan As Double = Math.Abs(.XAxis.Scale.Min - .XAxis.Scale.Max)
        Dim YSpan As Double = Math.Abs(.YAxis.Scale.Min - .YAxis.Scale.Max)
        If XSpan > YSpan Then
          Dim YSpanAdjust As Double = (XSpan - YSpan) / 2
          .YAxis.Scale.Min -= YSpanAdjust : .YAxis.Scale.Max += YSpanAdjust
        Else
          Dim XSpanAdjust As Double = (YSpan - XSpan) / 2
          .XAxis.Scale.Min -= XSpanAdjust : .XAxis.Scale.Max += XSpanAdjust
        End If

      End If
    End With
    UpdateDisplay(False)
  End Sub

  Public Sub GridOnOff(ByVal MajorGrid As Boolean, ByVal MinorGrid As Boolean)
    With MainGraph.GraphPane
      .XAxis.MajorGrid.IsVisible = MajorGrid
      .XAxis.MinorGrid.IsVisible = MinorGrid
      .YAxis.MajorGrid.IsVisible = MajorGrid
      .YAxis.MinorGrid.IsVisible = MinorGrid
    End With
  End Sub

  Public Sub SetCaptions(ByVal Title As String, ByVal XAxis As String, ByVal YAxis As String)
    SetCaptions(Title, XAxis, YAxis, String.Empty)
  End Sub

  Public Sub SetCaptions(ByVal Title As String, ByVal XAxis As String, ByVal YAxis As String, ByVal Y2Axis As String)
    MainGraph.GraphPane.Title.Text = Title
    MainGraph.GraphPane.XAxis.Title.Text = XAxis
    MainGraph.GraphPane.YAxis.Title.Text = YAxis
    If String.IsNullOrEmpty(Y2Axis) = False Then MainGraph.GraphPane.Y2Axis.Title.Text = Y2Axis
  End Sub

  Public Sub MaximizePlotArea()
    With MainGraph
      .GraphPane.Margin.All = 0
      .GraphPane.XAxis.Title.FontSpec.Size = 12
      .GraphPane.YAxis.Title.FontSpec.Size = 12
      .GraphPane.TitleGap = 0
      .GraphPane.Title.FontSpec.Size = 14
      .GraphPane.Legend.FontSpec.Size = 10
    End With
  End Sub

  Public Sub FormatLabelsSmaller()
    With MainGraph
      .GraphPane.Title.FontSpec.Size = 12
      .GraphPane.XAxis.Title.FontSpec.Size = 12
      .GraphPane.XAxis.Scale.FontSpec.Size = 8
      .GraphPane.YAxis.Title.FontSpec.Size = 12
      .GraphPane.YAxis.Scale.FontSpec.Size = 8
      .GraphPane.Legend.FontSpec.Size = 8
    End With
  End Sub

  Public Sub AddAnnotation(ByVal Text As String)
    Dim Annotate As New ZedGraph.TextObj(Text, 0.02, 0.10000000000000001, ZedGraph.CoordType.ChartFraction)
    Annotate.Location.AlignH = ZedGraph.AlignH.Left
    Annotate.Location.AlignV = ZedGraph.AlignV.Bottom
    Annotate.FontSpec.Size = 8
    Annotate.FontSpec.Border.IsVisible = True
    Annotate.FontSpec.Fill = New ZedGraph.Fill(Drawing.Color.LightGray)
    Annotate.FontSpec.StringAlignment = Drawing.StringAlignment.Near
    MainGraph.GraphPane.GraphObjList.Add(Annotate)
  End Sub

  Public Sub MarkXArea(ByVal XStart As Double, ByVal XStop As Double)
    Dim PvTDispMax As Double = MainGraph.GraphPane.YAxis.Scale.Max
    Dim PvTDispMin As Double = MainGraph.GraphPane.YAxis.Scale.Min
    Dim box As ZedGraph.BoxObj = New ZedGraph.BoxObj(XStart, PvTDispMax, XStop - XStart, PvTDispMin - PvTDispMax, Drawing.Color.Empty, Drawing.Color.Blue)
    box.Location.CoordinateFrame = ZedGraph.CoordType.AxisXYScale
    box.Location.AlignH = ZedGraph.AlignH.Left
    box.Location.AlignV = ZedGraph.AlignV.Top
    box.ZOrder = ZedGraph.ZOrder.E_BehindCurves
    MainGraph.GraphPane.GraphObjList.Clear()
    MainGraph.GraphPane.GraphObjList.Add(box)
  End Sub

    '''<summary>Plot constellation marker.</summary>
    '''<param name="Name">Name of the constellation.</param>
    '''<param name="X">X points of the markers.</param>
    '''<param name="Y">Y points of the markers.</param>
    '''<param name="Radius">Radius of the markers</param>
    Public Sub PlotConstellationMarker(ByVal Name As String, ByRef X As Single(), ByRef Y As Single(), ByVal Radius As Double, ByVal CircleColor As Drawing.Color)

        Dim Points_X As New List(Of Double)
        Dim Points_Y As New List(Of Double)
        Const PointsPerCircle As Integer = 100

        For Idx As Integer = 0 To X.GetUpperBound(0)
            'Draw circle
            For Angle As Double = 0 To 2 * Math.PI Step (2 * Math.PI / PointsPerCircle)
                Points_X.Add(X(Idx) + (Math.Cos(Angle) * Radius))
                Points_Y.Add(Y(Idx) + (Math.Sin(Angle) * Radius))
            Next Angle
            'Ensure to close circle
            Points_X.Add(X(Idx) + Radius)
            Points_Y.Add(Y(Idx))
            'Next circle
            Points_X.Add(Double.NaN)
            Points_Y.Add(Double.NaN)
        Next Idx

        Dim PlotStyle As New sGraphStyle(CircleColor, eCurveMode.Lines)
        PlotXvsY(Name, Points_X.ToArray, Points_Y.ToArray, PlotStyle)

        ' Rescale Plot because constellation markers might exceed the scaling of the
        ' actual data points.
        AutoScaleAxis()

    End Sub

    '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
    '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
    '''<param name="X">Vector of X axis values.</param>
    '''<param name="Y">Vector of Y axis values.</param>
    '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
    Public Sub PlotXvsY(ByRef CurveName As String, ByRef Elements As Dictionary(Of UInt32, UInteger), ByVal Style As sGraphStyle)
        Dim X(Elements.Count - 1) As Double
        Dim Y(Elements.Count - 1) As Double
        Dim Ptr As Integer = 0
        For Each Element As UInt32 In Elements.Keys
            X(Ptr) = Element
            Y(Ptr) = Elements(Element)
            Ptr += 1
        Next Element
        PlotXvsY(CurveName, X, Y, Style, False)
    End Sub

    '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
    '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
    '''<param name="X">Vector of X axis values.</param>
    '''<param name="Y">Vector of Y axis values.</param>
    '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
    Public Sub PlotXvsY(ByRef CurveName As String, ByRef X() As Double, ByRef Y() As Double, ByVal Style As sGraphStyle)
    PlotXvsY(CurveName, X, Y, Style, False)
  End Sub

  '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
  '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
  '''<param name="X">Vector of X axis values.</param>
  '''<param name="Y">Vector of Y axis values.</param>
  '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
  '''<param name="XMin">X axis, minimum value; if NaN is passed, the value is calculated automatically.</param>
  '''<param name="XMax">X axis, maximum value; if NaN is passed, the value is calculated automatically.</param>
  '''<param name="YMin">Y axis, minimum value; if NaN is passed, the value is calculated automatically.</param>
  '''<param name="YMax">Y axis, maximum value; if NaN is passed, the value is calculated automatically.</param>
  Public Sub PlotXvsY(ByRef CurveName As String, ByRef X() As Double, ByRef Y() As Double, ByVal Style As sGraphStyle, ByVal XMin As Double, ByVal XMax As Double, ByVal YMin As Double, ByVal YMax As Double)
    PlotXvsY(CurveName, X, Y, Style, False)
  End Sub

  '''<summary>Plot X vs Y data. This is the root plot routine.</summary>
  '''<param name="CurveName">Name of the curve. If the name already exists, only the data will be updated, but no new curve is build.</param>
  '''<param name="X">Vector of X axis values.</param>
  '''<param name="Y">Vector of Y axis values.</param>
  '''<param name="Style">Style to use (line, point, line and points, color, ...).</param>
  '''<param name="Use2ndYAxis">Use the 2nd Y axis for plotting?</param>
  Public Sub PlotXvsY(ByRef CurveName As String, ByRef X() As Double, ByRef Y() As Double, ByVal Style As sGraphStyle, ByVal Use2ndYAxis As Boolean)

    If CurveList.ContainsKey(CurveName) = False Then
      CurveList.Add(CurveName, MainGraph.GraphPane.AddCurve(CurveName, X, Y, Style.LineColor, Style.DotStyle))
    Else
      UpdateXvsY(CurveName, X, Y)
    End If

    With CurveList(CurveName)

      .Tag = CurveName

            'Asign to axis
            .IsY2Axis = Use2ndYAxis
            MainGraph.GraphPane.Y2Axis.IsVisible = Use2ndYAxis

            Style.ConfigureLineItem(CurveList(CurveName))

    End With

    MainGraph.AxisChange()

  End Sub

  Private Sub UpdateXvsY(ByRef CurveName As String, ByRef X As Double(), ByRef Y As Double())

    'If the curve name is not known, don't update
    If CurveList.ContainsKey(CurveName) = False Then Exit Sub

    'Process
    With CurveList(CurveName)
      .Clear()
      'Protect against corrupt X and Y data
      If (IsNothing(X) = False) And (IsNothing(Y) = False) Then
        If X.Length = Y.Length Then
          For Idx As Integer = 0 To X.GetUpperBound(0)
            .AddPoint(X(Idx), Y(Idx))
          Next
        End If
      End If

    End With
  End Sub

  Public Sub ShowLineAndDots(ByVal ShowLine As Boolean, ByVal ShowDots As Boolean)

    For Each Curve As String In CurveList.Keys
      With CurveList(Curve)
        .Line.IsVisible = ShowLine
        .Symbol.IsVisible = ShowDots
      End With
    Next Curve
    UpdateDisplay(False)

  End Sub

  Public Sub SaveAllTraces(ByVal FileName As String)

    Dim Columns As New List(Of String)

    Dim CurveCount As Integer = 0
    For Each CurveName As String In CurveList.Keys
      With CurveList(CurveName)
        For Idx As Integer = 0 To .Points.Count - 1
          If Columns.Count <= Idx Then
            Columns.Add(New String(CChar(";"), 2 * CurveCount))
          End If
          Columns(Idx) = Columns(Idx) & CurveList(CurveName).Points.Item(Idx).X.ToString.Trim & ";" & CurveList(CurveName).Points.Item(Idx).Y.ToString.Trim & ";"
        Next Idx
      End With
      CurveCount += 1
    Next CurveName

    IO.File.WriteAllText(FileName, Join(Columns.ToArray, Environment.NewLine))

  End Sub

End Class