Option Explicit On
Option Strict On

'''<summary>Class to host function for the statistics graph and text display.</summary>
Public Class cAstroStatDisp

    '''<summary>Properties.</summary>
    Public Prop As New cProp

    '''<summary>ZED graph plotter.</summary>
    Public Plotter As cZEDGraphService

    Public Class cProp

        <ComponentModel.Category("1. Generic")>
        <ComponentModel.DisplayName("   d) Bayer pattern")>
        <ComponentModel.Description("Bayer pattern")>
        <ComponentModel.DefaultValue("RGGB")>
        Public Property BayerPattern As String = "RGGB"

        <ComponentModel.Category("2. Plot")>
        <ComponentModel.DisplayName("   a) Plot single statistics")>
        <ComponentModel.DefaultValue(True)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.BooleanPropertyConverter_YesNo))>
        Public Property PlotSingleStatistics As Boolean = True

        <ComponentModel.Category("2. Plot")>
        <ComponentModel.DisplayName("   b) Plot mean statistics")>
        <ComponentModel.DefaultValue(True)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.BooleanPropertyConverter_YesNo))>
        Public Property PlotMeanStatistics As Boolean = True

        <ComponentModel.Category("2. Plot")>
        <ComponentModel.DisplayName("   c) Plot statistics - Mono")>
        <ComponentModel.DefaultValue(True)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.BooleanPropertyConverter_YesNo))>
        Public Property PlotStatisticsMono As Boolean = True

        <ComponentModel.Category("2. Plot")>
        <ComponentModel.DisplayName("   d) Plot limits fixed")>
        <ComponentModel.Description("True to auto-scale on min and max ADU, false to scale on data min and max")>
        <ComponentModel.DefaultValue(eXAxisScalingMode.Auto)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.EnumDesciptionConverter))>
        Public Property PlotLimitMode As eXAxisScalingMode = eXAxisScalingMode.Auto

        <ComponentModel.Category("2. Plot")>
        <ComponentModel.DisplayName("   e) Curve mode")>
        <ComponentModel.Description("True to auto-scale on min and max ADU, false to scale on data min and max")>
        <ComponentModel.DefaultValue(cZEDGraphService.eCurveMode.LinesAndPoints)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.EnumDesciptionConverter))>
        Public Property CurveMode As cZEDGraphService.eCurveMode = cZEDGraphService.eCurveMode.LinesAndPoints

        <ComponentModel.Category("3. Text")>
        <ComponentModel.DisplayName("   a) Single statistics log")>
        <ComponentModel.Description("Clear statistics log on every measurement")>
        <ComponentModel.DefaultValue(True)>
        <ComponentModel.TypeConverter(GetType(ComponentModelEx.BooleanPropertyConverter_YesNo))>
        Public Property Log_ClearStat As Boolean = True

        '''<summary>Get the channel name of the bayer pattern index.</summary>
        '''<param name="Idx">0-based index.</param>
        '''<returns>Channel name - if there are more channels with the same letter a number is added beginning with the 2nd channel.</returns>
        Public Function BayerPatternName(ByVal PatIdx As Integer) As String
            If PatIdx > BayerPattern.Length - 1 Then Return "?"
            Dim Dict As New Dictionary(Of String, Integer)
            Dim ColorName As String = String.Empty
            For Idx As Integer = 0 To PatIdx
                ColorName = BayerPattern.Substring(Idx, 1)
                If Dict.ContainsKey(ColorName) = False Then
                    Dict.Add(ColorName, 0)
                Else
                    Dict(ColorName) += 1
                End If
            Next Idx
            If Dict(ColorName) > 0 Then
                Return ColorName & Dict(ColorName).ValRegIndep
            Else
                Return ColorName
            End If
        End Function

        '''<summary>Get the channel name of all bayer pattern index.</summary>
        '''<param name="Idx">0-based index.</param>
        '''<returns>Channel name.</returns>
        Public Function BayerPatternNames() As List(Of String)
            Dim RetVal As New List(Of String)
            For Idx As Integer = 0 To BayerPattern.Length - 1
                RetVal.Add(BayerPatternName(Idx))
            Next Idx
            Return RetVal
        End Function

    End Class

    Public Sub Plot(ByVal CaptureCount As Int32, ByRef SingleStat As AstroNET.Statistics.sStatistics, ByRef LoopStat As AstroNET.Statistics.sStatistics)

        Dim CurrentCurveWidth As Integer = 1
        Dim MeanCurveWidth As Integer = 2
        'If IsNothing(Plotter) = True Then Plotter = New cZEDGraphService(zgcMain)
        Plotter.Clear()
        If Prop.PlotMeanStatistics = True Or Prop.PlotSingleStatistics = True Then
            'Mean statistics
            If CaptureCount > 1 And LoopStat.Count > 1 And Prop.PlotMeanStatistics = True Then
                If IsNothing(LoopStat.BayerHistograms_Int) = False Then
                    Plotter.PlotXvsY(Prop.BayerPatternName(0) & "[0,0] mean", LoopStat.BayerHistograms_Int(0, 0), LoopStat.Count, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Red, Prop.CurveMode, MeanCurveWidth))
                    Plotter.PlotXvsY(Prop.BayerPatternName(0) & "[0,1] mean", LoopStat.BayerHistograms_Int(0, 1), LoopStat.Count, New cZEDGraphService.sGraphStyle(System.Drawing.Color.LightGreen, Prop.CurveMode, MeanCurveWidth))
                    Plotter.PlotXvsY(Prop.BayerPatternName(0) & "[1,0] mean", LoopStat.BayerHistograms_Int(1, 0), LoopStat.Count, New cZEDGraphService.sGraphStyle(System.Drawing.Color.DarkGreen, Prop.CurveMode, MeanCurveWidth))
                    Plotter.PlotXvsY(Prop.BayerPatternName(0) & "[1,1] mean", LoopStat.BayerHistograms_Int(1, 1), LoopStat.Count, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Blue, Prop.CurveMode, MeanCurveWidth))
                End If
                If IsNothing(LoopStat.MonochromHistogram_Int) = False And Prop.PlotStatisticsMono = True Then
                    Plotter.PlotXvsY("Mono mean", LoopStat.MonochromHistogram_Int, LoopStat.Count, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Black, Prop.CurveMode, MeanCurveWidth))
                End If
            End If
            'Current statistics
            If Prop.PlotSingleStatistics = True And IsNothing(SingleStat.BayerHistograms_Int) = False Then
                Plotter.PlotXvsY(Prop.BayerPatternName(0) & "[0,0]", SingleStat.BayerHistograms_Int(0, 0), 1, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Red, Prop.CurveMode, CurrentCurveWidth))
                Plotter.PlotXvsY(Prop.BayerPatternName(1) & "[0,1]", SingleStat.BayerHistograms_Int(0, 1), 1, New cZEDGraphService.sGraphStyle(System.Drawing.Color.LightGreen, Prop.CurveMode, CurrentCurveWidth))
                Plotter.PlotXvsY(Prop.BayerPatternName(2) & "[1,0]", SingleStat.BayerHistograms_Int(1, 0), 1, New cZEDGraphService.sGraphStyle(System.Drawing.Color.DarkGreen, Prop.CurveMode, CurrentCurveWidth))
                Plotter.PlotXvsY(Prop.BayerPatternName(3) & "[1,1]", SingleStat.BayerHistograms_Int(1, 1), 1, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Blue, Prop.CurveMode, CurrentCurveWidth))
                If IsNothing(SingleStat.MonochromHistogram_Int) = False And Prop.PlotStatisticsMono = True Then
                    Plotter.PlotXvsY("Mono", SingleStat.MonochromHistogram_Int, 1, New cZEDGraphService.sGraphStyle(System.Drawing.Color.Black, Prop.CurveMode, CurrentCurveWidth))
                End If
            End If
            Select Case Prop.PlotLimitMode
                Case eXAxisScalingMode.Auto
                    Plotter.ManuallyScaleXAxis(LoopStat.MonoStatistics_Int.Min.Key, LoopStat.MonoStatistics_Int.Max.Key)
                Case eXAxisScalingMode.FullRange16Bit
                    Plotter.ManuallyScaleXAxis(0, 65536)
                Case eXAxisScalingMode.LeaveAsIs
                    'Just do nothing ...
            End Select

            Plotter.AutoScaleYAxisLog()
            Plotter.GridOnOff(True, True)
            Plotter.ForceUpdate()
        End If

    End Sub

End Class