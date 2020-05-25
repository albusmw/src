Option Explicit On
Option Strict On

'''<summary>Type for results of ADU value counts.</summary>
Imports ADUCount = System.UInt64
'''<summary>Type for interger ADU statistics.</summary>
Imports ADUFixed = System.Int64

'This class is the container for all astronomical statistics functions
'All other modules with statistics should not be used any more and put their code here
'Old modules:
' - C:\GIT\src\atomics\cStatistics.vb
Namespace AstroNET

    Public Class Statistics

        Public Enum eDataMode
            [Invalid] = 1
            [UInt16] = 1
            [UInt32] = 2
            [Int32] = 3
            [Float32] = 4
        End Enum

        '''<summary>Instance of Intel IPP library call.</summary>
        Private IntelIPP As cIntelIPP = Nothing

        '''<summary>Image data in UInt16 mode.</summary>
        Public DataProcessor_UInt16 As New cStatMultiThread_UInt16

        '''<summary>Image data in Int32 mode.</summary>
        Public DataProcessor_Int32 As New cStatMultiThread_Int32

        '''<summary>Image data in UInt32 mode.</summary>
        Public DataProcessor_UInt32 As New cStatMultiThread_UInt32

        '''<summary>Image data in Float32 mode.</summary>
        Public DataProcessor_Float32 As New cStatMultiThread_Float32

        Private Const UInt64_1 As UInt64 = CType(1, UInt64)

        Public Sub ResetAllProcessors()
            Reset_UInt16()
            Reset_UInt32()
            Reset_Float32()
            DataProcessor_Int32.ImageData = {{}}
        End Sub

        Public Sub Reset_UInt16()
            For Idx As Integer = 0 To 3
                DataProcessor_UInt16.ImageData(Idx).Data = {}
            Next Idx
        End Sub

        Public Sub Reset_UInt32()
            For Idx As Integer = 0 To 3
                DataProcessor_UInt32.ImageData(Idx).Data = {}
            Next Idx
        End Sub

        Public Sub Reset_Float32()
            For Idx As Integer = 0 To 3
                DataProcessor_Float32.ImageData(Idx).Data = {}
            Next Idx
        End Sub

        '''<summary>Return which data type is currenty loaded (only 1 data type can be loaded).</summary>
        Public ReadOnly Property DataModeType() As sStatistics.eDataMode
            Get
                Select Case DataMode
                    Case eDataMode.Float32
                        Return sStatistics.eDataMode.Float
                    Case eDataMode.UInt16, eDataMode.UInt32, eDataMode.Int32
                        Return sStatistics.eDataMode.Fixed
                    Case Else
                        Return sStatistics.eDataMode.Invalid
                End Select
                Return Nothing
            End Get
        End Property


        '''<summary>Return which data type is currenty loaded (only 1 data type can be loaded).</summary>
        Public ReadOnly Property DataMode() As eDataMode
            Get
                If DataProcessor_UInt16.ImageData(0).Data.LongLength > 0 Then Return eDataMode.UInt16
                If DataProcessor_UInt32.ImageData(0).Data.LongLength > 0 Then Return eDataMode.UInt32
                If DataProcessor_Int32.ImageData.LongLength > 0 Then Return eDataMode.Int32
                If DataProcessor_Float32.ImageData(0).Data.LongLength > 0 Then Return eDataMode.Float32
                Return Nothing
            End Get
        End Property

        '''<summary>Return the current dimensions of the loaded data.</summary>
        Public ReadOnly Property Dimensions() As String
            Get
                If DataProcessor_UInt16.ImageData(0).Data.LongLength > 0 Then
                    Return DataProcessor_UInt16.ImageData(0).NAXIS1.ValRegIndep & "x" & DataProcessor_UInt16.ImageData(0).NAXIS2.ValRegIndep
                End If
                If DataProcessor_UInt32.ImageData(0).Data.LongLength > 0 Then
                    Return DataProcessor_UInt32.ImageData(0).NAXIS1.ValRegIndep & "x" & DataProcessor_UInt32.ImageData(0).NAXIS2.ValRegIndep
                End If
                If DataProcessor_Float32.ImageData(0).Data.LongLength > 0 Then
                    Return DataProcessor_Float32.ImageData(0).NAXIS1.ValRegIndep & "x" & DataProcessor_Float32.ImageData(0).NAXIS2.ValRegIndep
                End If
                If DataProcessor_Int32.ImageData.LongLength > 0 Then
                    Return (DataProcessor_Int32.ImageData.GetUpperBound(0) + 1).ValRegIndep & "x" & (DataProcessor_Int32.ImageData.GetUpperBound(1) + 1).ValRegIndep
                End If
                Return Nothing
            End Get
        End Property

        '''<summary>Constructor that creates an Intel IPP reference.</summary>
        '''<param name="IPPPath">Path to ipps.dll and ippvm.dll - if not set IPP will not be used.</param>
        Public Sub New(ByVal IPPPath As String)
            IntelIPP = New cIntelIPP(IPPPath)
        End Sub

        '''<summary>Constructor that creates an Intel IPP reference.</summary>
        '''<param name="ExistingIntelIPP">Reference to an existing Intel IPP class library.</param>
        Public Sub New(ByRef ExistingIntelIPP As cIntelIPP)
            IntelIPP = ExistingIntelIPP
        End Sub

        '''<summary>Total statistics - available as per-channel bayer statistics and as combined statistics.</summary>
        Public Structure sStatistics

            Public Enum eDataMode
                [Invalid] = 0
                [Fixed] = 1
                [Float] = 2
            End Enum

            '''<summary>Full-resolution histogram data - bayer data.</summary>
            Public BayerHistograms_Int(,) As Dictionary(Of ADUFixed, ADUCount)
            '''<summary>Full-resolution histogram data - mono data, sorted.</summary>
            Public MonochromHistogram_Int As Dictionary(Of ADUFixed, ADUCount)
            '''<summary>Statistics for each channel.</summary>
            Public BayerStatistics_Int(,) As sSingleChannelStatistics_Int
            '''<summary>Statistics for each channel.</summary>
            Public MonoStatistics_Int As sSingleChannelStatistics_Int

            '''<summary>Full-resolution histogram data - bayer data.</summary>
            Public BayerHistograms_Float32(,) As Dictionary(Of Single, ADUCount)
            '''<summary>Full-resolution histogram data - mono data, sorted.</summary>
            Public MonochromHistogram_Float32 As Dictionary(Of Single, ADUCount)
            '''<summary>Statistics for each channel.</summary>
            Public BayerStatistics_Float32(,) As sSingleChannelStatistics_Float32
            '''<summary>Statistics for each channel.</summary>
            Public MonoStatistics_Float32 As sSingleChannelStatistics_Float32

            Public Function BayerHistograms_Int_Present(ByVal Idx0 As Integer, ByVal Idx1 As Integer, ByVal Key As ADUFixed) As Boolean
                If IsNothing(BayerHistograms_Int) = True Then Return False
                If IsNothing(BayerHistograms_Int(Idx0, Idx1)) = True Then Return False
                If BayerHistograms_Int(Idx0, Idx1).ContainsKey(Key) = False Then Return False
                Return True
            End Function

            '''<summary>Which data are present in the statistics.</summary>
            Public ReadOnly Property DataMode() As eDataMode
                Get
                    If IsNothing(BayerHistograms_Float32) = False Then Return eDataMode.Float
                    If IsNothing(BayerHistograms_Int) = False Then Return eDataMode.Fixed
                    Return eDataMode.Invalid
                End Get
            End Property

            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport(ByVal ChannelNames As List(Of String)) As List(Of String)
                Select Case DataMode
                    Case eDataMode.Float
                        Return StatisticsReport_Float32(ChannelNames, String.Empty)
                    Case eDataMode.Fixed
                        Return StatisticsReport_Int(ChannelNames, String.Empty)
                    Case Else
                        Return New List(Of String)
                End Select
            End Function

            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport_Int(ByVal ChannelNames As List(Of String), ByVal Indent As String) As List(Of String)
                Dim RetVal As New List(Of String)
                RetVal.Add(Indent & "Property".PadRight(sSingleChannelStatistics_Int.ReportHeaderLength) & ": " & "Mono".PadRight(sSingleChannelStatistics_Int.ReportValueLength) & "|")
                For Each Entry As String In MonoStatistics_Int.StatisticsReport
                    RetVal.Add(Indent & "  " & Entry & "|")
                Next Entry
                Dim ChannelIdx As Integer = 0
                For Idx1 As Integer = 0 To BayerStatistics_Int.GetUpperBound(0)
                    For Idx2 As Integer = 0 To BayerStatistics_Int.GetUpperBound(1)
                        RetVal(0) &= (ChannelNames(ChannelIdx) & "[" & Idx1.ValRegIndep & ":" & Idx2.ValRegIndep & "]").PadRight(sSingleChannelStatistics_Int.ReportValueLength) & "|"
                        Dim LineIdx As Integer = 1
                        For Each Entry As String In BayerStatistics_Int(Idx1, Idx2).StatisticsReport
                            RetVal(LineIdx) &= Entry.Substring(sSingleChannelStatistics_Int.ReportHeaderLength) & "|"
                            LineIdx += 1
                        Next Entry
                        ChannelIdx += 1
                    Next Idx2
                Next Idx1
                Return RetVal
            End Function

            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport_Float32(ByVal ChannelNames As List(Of String), ByVal Indent As String) As List(Of String)
                Dim RetVal As New List(Of String)
                RetVal.Add(Indent & "Property".PadRight(sSingleChannelStatistics_Float32.ReportHeaderLength) & ": " & "Mono".PadRight(sSingleChannelStatistics_Float32.ReportValueLength) & "|")
                For Each Entry As String In MonoStatistics_Float32.StatisticsReport
                    RetVal.Add(Indent & "  " & Entry & "|")
                Next Entry
                Dim ChannelIdx As Integer = 0
                For Idx1 As Integer = 0 To BayerStatistics_Float32.GetUpperBound(0)
                    For Idx2 As Integer = 0 To BayerStatistics_Float32.GetUpperBound(1)
                        RetVal(0) &= (ChannelNames(ChannelIdx) & "[" & Idx1.ValRegIndep & ":" & Idx2.ValRegIndep & "]").PadRight(sSingleChannelStatistics_Float32.ReportValueLength) & "|"
                        Dim LineIdx As Integer = 1
                        For Each Entry As String In BayerStatistics_Float32(Idx1, Idx2).StatisticsReport
                            RetVal(LineIdx) &= Entry.Substring(sSingleChannelStatistics_Float32.ReportHeaderLength) & "|"
                            LineIdx += 1
                        Next Entry
                        ChannelIdx += 1
                    Next Idx2
                Next Idx1
                Return RetVal
            End Function

            '''<summary>Returns the number of values that are above the given value.</summary>
            Public Function BayerHistograms_Int_ValuesAbove(ByVal Idx0 As Integer, ByVal Idx1 As Integer, ByVal X As ADUFixed) As ADUCount
                Dim RetVal As ADUCount = 0
                For Each Key As ADUFixed In BayerHistograms_Int(Idx0, Idx1).Keys
                    If Key > X Then RetVAl += BayerHistograms_Int(Idx0, Idx1)(Key)
                Next Key
                Return RetVAl
            End Function

            '''<summary>Returns the number of values that are above the given value.</summary>
            Public Function MonochromHistogram_Int_ValuesAbove(ByVal X As ADUFixed) As ADUCount
                Dim RetVal As ADUCount = 0
                For Each Key As ADUFixed In MonochromHistogram_Int.Keys
                    If Key > X Then RetVal += MonochromHistogram_Int(Key)
                Next Key
                Return RetVal
            End Function


        End Structure

        '''<summary>Statistic information of one channel (RGB or total).</summary>
        '''<remarks>The maximum word with is taken as pixel values to cover all fixed-point formats ...</remarks>
        Public Structure sSingleChannelStatistics_Int
            '''<summary>Number of characters in the header of the report.</summary>
            Public Shared ReadOnly Property ReportHeaderLength As Integer = 20
            '''<summary>Number of characters in the value of the report.</summary>
            Public Shared ReadOnly Property ReportValueLength As Integer = 16
            '''<summary>Width [pixel] of the last image.</summary>
            Public Width As UInt32
            '''<summary>Height [pixel] of the last image.</summary>
            Public Height As UInt32
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As UInt64
            '''<summary>Maximum value occured (value and number of pixel that have this value).</summary>
            Public Max As KeyValuePair(Of ADUFixed, ADUCount)
            '''<summary>Minimum value occured (value and number of pixel that have this value).</summary>
            Public Min As KeyValuePair(Of ADUFixed, ADUCount)
            '''<summary>Value where half of the samples are below and half are above.</summary>
            Public Median As Int64
            '''<summary>Arithmetic mean value.</summary>
            Public Mean As Double
            '''<summary>Mean value of squared values.</summary>
            Public MeanPow2 As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public StdDev As Double
            '''<summary>Number of different ADU values in the data.</summary>
            Public DifferentADUValues As Integer
            '''<summary>Number of different ADU values in 25-75 pct range.</summary>
            Public ADUValues2575 As Integer
            '''<summary>Distance between the histogram X axis points.</summary>
            Public HistXDist As Dictionary(Of Long, UInt64)
            '''<summary>Percentile.</summary>
            Public Percentile As Dictionary(Of Integer, Int64)
            '''<summary>Pixel value that is present the most often.</summary>
            Public Modus As KeyValuePair(Of ADUFixed, ADUCount)
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public ReadOnly Property Variance As Double
                Get
                    Return StdDev ^ 2
                End Get
            End Property
            '''<summary>Init all inner variables.</summary>
            Public Shared Function InitForShort() As sSingleChannelStatistics_Int
                Dim RetVal As New sSingleChannelStatistics_Int
                RetVal.Width = 0
                RetVal.Height = 0
                RetVal.Samples = 0
                RetVal.Max = Nothing
                RetVal.Min = Nothing
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentADUValues = 0
                RetVal.HistXDist = New Dictionary(Of Long, UInt64)
                RetVal.Median = Int64.MinValue
                RetVal.Percentile = New Dictionary(Of Integer, Int64)
                RetVal.Modus = Nothing
                Return RetVal
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            '''<param name="DispHeader">TRUE to display the header, FALSE else.</param>
            Public Function StatisticsReport() As List(Of String)
                Dim NotPresent As String = New String("-"c, ReportValueLength)
                Dim RetVal As New List(Of String)
                Dim HistXDist_keys As List(Of Long) = HistXDist.KeyList : If IsNothing(HistXDist_keys) = True Then HistXDist_keys = New List(Of Long)
                Dim TotalPixel As String = ((Samples / 1000000).ValRegIndep("0.0") & "M")
                If Samples < 1000000 Then TotalPixel = ((Samples / 1000).ValRegIndep("0.0") & "k")
                RetVal.Add("Dimensions        : " & (Width.ValRegIndep & "x" & Height.ValRegIndep).PadLeft(ReportValueLength))
                RetVal.Add("Total pixel       : " & Samples.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Total pixel       : " & TotalPixel.PadLeft(ReportValueLength))
                RetVal.Add("ADU values count  : " & DifferentADUValues.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("  in 25-75 pct    : " & ADUValues2575.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Min value         : " & (Min.Key.ValRegIndep & " (" & Min.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Modus value       : " & (Modus.Key.ValRegIndep & " (" & Modus.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Max value         : " & (Max.Key.ValRegIndep & " (" & Max.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Median value      : " & Median.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Mean value        : " & Format(Mean, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Standard dev.     : " & Format(StdDev, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Variance          : " & Format(Variance, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                'Data on histogram of ADU stepping
                If HistXDist_keys.Count = 0 Then
                    RetVal.Add("ADU step size min : " & NotPresent.PadLeft(ReportValueLength))
                    RetVal.Add("ADU different step: " & NotPresent.PadLeft(ReportValueLength))
                Else
                    RetVal.Add("ADU step size min : " & Format(HistXDist_keys(0), "####0").ToString.Trim.PadLeft(ReportValueLength))
                    RetVal.Add("ADU different step: " & Format(HistXDist_keys.Count, "####0").ToString.Trim.PadLeft(ReportValueLength))
                End If
                'Percentile report
                For Each Pct As Integer In New Integer() {1, 5, 10, 25, 50, 75, 90, 95, 99}
                    If IsNothing(Percentile) = False Then
                        If Percentile.ContainsKey(Pct) Then RetVal.Add(("Percentil - " & Pct.ToString.Trim.PadLeft(2) & " %  : ").PadRight(ReportHeaderLength) & Format(Percentile(Pct)).ToString.Trim.PadLeft(ReportValueLength))
                    End If
                Next Pct
                Return RetVal
            End Function
        End Structure

        '''<summary>Statistic information of one channel (RGB or total).</summary>
        '''<remarks>The maximum word with is taken as pixel values to cover all fixed-point formats ...</remarks>
        Public Structure sSingleChannelStatistics_Float32
            '''<summary>Number of characters in the header of the report.</summary>
            Public Shared ReadOnly Property ReportHeaderLength As Integer = 20
            '''<summary>Number of characters in the value of the report.</summary>
            Public Shared ReadOnly Property ReportValueLength As Integer = 16
            '''<summary>Width [pixel] of the last image.</summary>
            Public Width As UInt32
            '''<summary>Height [pixel] of the last image.</summary>
            Public Height As UInt32
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As ADUCount
            '''<summary>Maximum value occured (value and number of pixel that have this value).</summary>
            Public Max As KeyValuePair(Of Single, ADUCount)
            '''<summary>Minimum value occured (value and number of pixel that have this value).</summary>
            Public Min As KeyValuePair(Of Single, ADUCount)
            '''<summary>Value where half of the samples are below and half are above.</summary>
            Public Median As Single
            '''<summary>Arithmetic mean value.</summary>
            Public Mean As Double
            '''<summary>Mean value of squared values.</summary>
            Public MeanPow2 As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public StdDev As Double
            '''<summary>Number of different ADU values in the data.</summary>
            Public DifferentADUValues As Integer
            '''<summary>Number of different ADU values in 25-75 pct range.</summary>
            Public ADUValues2575 As Integer
            '''<summary>Distance between the histogram X axis points.</summary>
            Public HistXDist As Dictionary(Of Single, UInt64)
            '''<summary>Percentile.</summary>
            Public Percentile As Dictionary(Of Integer, Single)
            '''<summary>Pixel value that is present the most often.</summary>
            Public Modus As KeyValuePair(Of Single, ADUCount)
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public ReadOnly Property Variance As Double
                Get
                    Return StdDev ^ 2
                End Get
            End Property
            '''<summary>Init all inner variables.</summary>
            Public Shared Function Init() As sSingleChannelStatistics_Float32
                Dim RetVal As New sSingleChannelStatistics_Float32
                RetVal.Width = 0
                RetVal.Height = 0
                RetVal.Samples = 0
                RetVal.Max = Nothing
                RetVal.Min = Nothing
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentADUValues = 0
                RetVal.HistXDist = New Dictionary(Of Single, UInt64)
                RetVal.Median = Int64.MinValue
                RetVal.Percentile = New Dictionary(Of Integer, Single)
                RetVal.Modus = Nothing
                Return RetVal
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            '''<param name="DispHeader">TRUE to display the header, FALSE else.</param>
            Public Function StatisticsReport() As List(Of String)
                Dim NotPresent As String = New String("-"c, ReportValueLength)
                Dim RetVal As New List(Of String)
                Dim HistXDist_keys As List(Of Single) = HistXDist.KeyList
                Dim TotalPixel As String = ((Samples / 1000000).ValRegIndep("0.0") & "M")
                If Samples < 1000000 Then TotalPixel = ((Samples / 1000).ValRegIndep("0.0") & "k")
                RetVal.Add("Dimensions        : " & (Width.ValRegIndep & "x" & Height.ValRegIndep).PadLeft(ReportValueLength))
                RetVal.Add("Total pixel       : " & Samples.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Total pixel       : " & TotalPixel.PadLeft(ReportValueLength))
                RetVal.Add("ADU values count  : " & DifferentADUValues.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("  in 25-75 pct    : " & ADUValues2575.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Min value         : " & (Min.Key.ValRegIndep & " (" & Min.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Modus value       : " & (Modus.Key.ValRegIndep & " (" & Modus.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Max value         : " & (Max.Key.ValRegIndep & " (" & Max.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Median value      : " & Median.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Mean value        : " & Format(Mean, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Standard dev.     : " & Format(StdDev, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Variance          : " & Format(Variance, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                'Data on histogram of ADU stepping
                If HistXDist_keys.Count = 0 Then
                    RetVal.Add("ADU step size min : " & NotPresent.PadLeft(ReportValueLength))
                    RetVal.Add("ADU different step: " & NotPresent.PadLeft(ReportValueLength))
                Else
                    RetVal.Add("ADU step size min : " & Format(HistXDist_keys(0), "####0").ToString.Trim.PadLeft(ReportValueLength))
                    RetVal.Add("ADU different step: " & Format(HistXDist_keys.Count, "####0").ToString.Trim.PadLeft(ReportValueLength))
                End If
                'Percentile report
                For Each Pct As Integer In New Integer() {1, 5, 10, 25, 50, 75, 90, 95, 99}
                    If Percentile.ContainsKey(Pct) Then RetVal.Add(("Percentil - " & Pct.ToString.Trim.PadLeft(2) & " %  : ").PadRight(ReportHeaderLength) & Format(Percentile(Pct)).ToString.Trim.PadLeft(ReportValueLength))
                Next Pct
                Return RetVal
            End Function
        End Structure

        '''<summary>Calculate the image statistics of the passed image data.</summary>
        Public Function ImageStatistics(ByVal DataMode As AstroNET.Statistics.sStatistics.eDataMode) As sStatistics

            'Calculate statistics for "first pane"
            Dim RetVal As New sStatistics
            Select Case DataMode
                Case sStatistics.eDataMode.Float
                    RetVal.BayerHistograms_Float32 = BayerStatistics_Float32()  'Calculate a 2x2 bayer statistics (also for mono data as thread-based will speed up ...)
                    CalculateAllFromBayerStatistics(DataMode, RetVal)           'Add all other data (mono histo and statistics)
                Case sStatistics.eDataMode.Fixed
                    RetVal.BayerHistograms_Int = BayerStatistics_Int(0)         'Calculate a 2x2 bayer statistics (also for mono data as thread-based will speed up ...)
                    CalculateAllFromBayerStatistics(DataMode, RetVal)           'Add all other data (mono histo and statistics)
                    If DataProcessor_UInt16.ImageData(1).Length > 1 Then
                        ClearBayerStatistics(RetVal)                            'Clear bayer statistics
                        SetReadColorStat(RetVal, RetVal, 0, 0)                  'Make 1st channel to red stats
                        Dim Green As New sStatistics                            'Prepare new statistics for green
                        Green.BayerHistograms_Int = BayerStatistics_Int(1)      'Calculate 2nd layer (Index=1 - green)
                        CalculateAllFromBayerStatistics(DataMode, Green)        'Add all other data (mono histo and statistics)
                        SetReadColorStat(RetVal, Green, 0, 1)                   'Make 1st channel to red stats
                        Dim Blue As New sStatistics                             'Prepare new statistics for blue
                        Blue.BayerHistograms_Int = BayerStatistics_Int(2)       'Calculate 3rd layer (Index=2 - blue)
                        CalculateAllFromBayerStatistics(DataMode, Blue)         'Add all other data (mono histo and statistics)
                        SetReadColorStat(RetVal, Blue, 1, 1)                    'Make 1st channel to red stats
                    End If
            End Select

            'Return results
            Return RetVal

        End Function

        '''<summary>Reset the bayer statistics to prepare for a real color statistics.</summary>
        Private Sub ClearBayerStatistics(ByRef Results As sStatistics)
            With Results
                .BayerHistograms_Int(0, 0) = Nothing
                .BayerHistograms_Int(1, 0) = Nothing
                .BayerHistograms_Int(0, 1) = Nothing
                .BayerHistograms_Int(1, 1) = Nothing
                .BayerStatistics_Int(0, 0) = Nothing
                .BayerStatistics_Int(1, 0) = Nothing
                .BayerStatistics_Int(0, 1) = Nothing
                .BayerStatistics_Int(1, 1) = Nothing
            End With
        End Sub

        '''<summary>Set a certain bayer channel.</summary>
        Private Sub SetReadColorStat(ByRef StatisticsToSet As sStatistics, ByVal NewStatistics As sStatistics, ByVal BIdx0 As Integer, ByVal BIdx1 As Integer)
            StatisticsToSet.BayerStatistics_Int(BIdx0, BIdx1) = NewStatistics.MonoStatistics_Int
            StatisticsToSet.BayerHistograms_Int(BIdx0, BIdx1) = NewStatistics.MonochromHistogram_Int
        End Sub

        '''<summary>Combine 2 SingleChannelStatistics elements (e.g. to calculate the aggregated statistic for multi-frame capture).</summary>
        Public Shared Function CombineStatistics(ByVal DataMode As AstroNET.Statistics.sStatistics.eDataMode, ByVal StatA As sStatistics, ByVal CombinedStatistics As sStatistics) As sStatistics
            Dim RetVal As New sStatistics
            '1.) Combine to 2 histograms
            If IsNothing(StatA.BayerHistograms_Int) = False Then
                ReDim RetVal.BayerHistograms_Int(StatA.BayerHistograms_Int.GetUpperBound(0), StatA.BayerHistograms_Int.GetUpperBound(1))
                For BayIdx1 As Integer = 0 To StatA.BayerHistograms_Int.GetUpperBound(0)
                    For BayIdx2 As Integer = 0 To StatA.BayerHistograms_Int.GetUpperBound(1)
                        RetVal.BayerHistograms_Int(BayIdx1, BayIdx2) = New Dictionary(Of Int64, UInt64)
                        'Init return bayer histogram with StatA data
                        For Each PixelValue As ADUFixed In StatA.BayerHistograms_Int(BayIdx1, BayIdx2).Keys
                            RetVal.BayerHistograms_Int(BayIdx1, BayIdx2).Add(PixelValue, StatA.BayerHistograms_Int(BayIdx1, BayIdx2)(PixelValue))
                        Next PixelValue
                        'Combine with StatB data
                        If IsNothing(CombinedStatistics.BayerHistograms_Int) = False Then
                            For Each PixelValue As ADUFixed In CombinedStatistics.BayerHistograms_Int(BayIdx1, BayIdx2).Keys
                                RetVal.BayerHistograms_Int(BayIdx1, BayIdx2).AddTo(PixelValue, CombinedStatistics.BayerHistograms_Int(BayIdx1, BayIdx2)(PixelValue))
                            Next PixelValue
                        End If
                        RetVal.BayerHistograms_Int(BayIdx1, BayIdx2) = RetVal.BayerHistograms_Int(BayIdx1, BayIdx2).SortDictionary
                    Next BayIdx2
                Next BayIdx1
                CalculateAllFromBayerStatistics(DataMode, RetVal)
            End If
            Return RetVal
        End Function

        '''<summary>Calculate all statistic data (mono histo and statistics) from the passed bayer statistics.</summary>
        Private Shared Sub CalculateAllFromBayerStatistics(ByVal DataMode As AstroNET.Statistics.sStatistics.eDataMode, ByRef RetVal As sStatistics)
            'Calculate a monochromatic statistics from the bayer histograms
            Select Case DataMode
                Case sStatistics.eDataMode.Float
                    RetVal.MonochromHistogram_Float32 = CombineBayerToMonoStatistics(RetVal.BayerHistograms_Float32)
                    ReDim RetVal.BayerStatistics_Float32(RetVal.BayerHistograms_Float32.GetUpperBound(0), RetVal.BayerHistograms_Float32.GetUpperBound(1))
                    For Idx1 As Integer = 0 To RetVal.BayerHistograms_Float32.GetUpperBound(0)
                        For Idx2 As Integer = 0 To RetVal.BayerHistograms_Float32.GetUpperBound(1)
                            RetVal.BayerStatistics_Float32(Idx1, Idx2) = CalcStatisticFromHistogram(RetVal.BayerHistograms_Float32(Idx1, Idx2))
                        Next Idx2
                    Next Idx1
                    RetVal.MonoStatistics_Float32 = CalcStatisticFromHistogram(RetVal.MonochromHistogram_Float32)
                Case sStatistics.eDataMode.Fixed
                    RetVal.MonochromHistogram_Int = CombineBayerToMonoStatistics(RetVal.BayerHistograms_Int)
                    ReDim RetVal.BayerStatistics_Int(RetVal.BayerHistograms_Int.GetUpperBound(0), RetVal.BayerHistograms_Int.GetUpperBound(1))
                    For Idx1 As Integer = 0 To RetVal.BayerHistograms_Int.GetUpperBound(0)
                        For Idx2 As Integer = 0 To RetVal.BayerHistograms_Int.GetUpperBound(1)
                            RetVal.BayerStatistics_Int(Idx1, Idx2) = CalcStatisticFromHistogram(RetVal.BayerHistograms_Int(Idx1, Idx2))
                        Next Idx2
                    Next Idx1
                    RetVal.MonoStatistics_Int = CalcStatisticFromHistogram(RetVal.MonochromHistogram_Int)
            End Select
        End Sub

        '''<summary>Calculate the statistic data from the passed histogram data.</summary>
        '''<param name="Histogram">Calculated histogram data.</param>
        Private Shared Function CalcStatisticFromHistogram(ByRef Histogram As Dictionary(Of ADUFixed, ADUCount)) As sSingleChannelStatistics_Int

            If IsNothing(Histogram) = True Then Return Nothing

            Dim RetVal As sSingleChannelStatistics_Int = sSingleChannelStatistics_Int.InitForShort()
            Dim AllADUValues As List(Of ADUFixed) = Histogram.KeyList
            AllADUValues.Sort()

            'Count number of samples
            For Each PixelValue As Int64 In Histogram.Keys
                RetVal.Samples += Histogram(PixelValue)
            Next PixelValue

            'Store number of different sample values
            RetVal.DifferentADUValues = Histogram.Count

            'Init statistics calculation
            Dim SamplesSeen As UInt64 = 0
            Dim MeanSum As Double = 0
            Dim MeanPow2Sum As Double = 0
            RetVal.Min = New KeyValuePair(Of Int64, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.Max = New KeyValuePair(Of Int64, UInt64)(AllADUValues(AllADUValues.Count - 1), Histogram(AllADUValues(AllADUValues.Count - 1)))
            RetVal.Modus = New KeyValuePair(Of Int64, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.HistXDist = New Dictionary(Of Long, UInt64)

            'Move over the histogram for normal statistics
            RetVal.ADUValues2575 = 0
            For Each ADUValue As ADUFixed In AllADUValues
                Dim ValueCount As ADUCount = Histogram(ADUValue)                                                                            'number of values with this ADU value
                SamplesSeen += ValueCount                                                                                                   'total pixel processed up to now
                Dim WeightCount As Double = (CType(ADUValue, Double) * CType(ValueCount, Double))                                           'ADUValue^2
                Dim WeightPow2 As Double = (CType(ADUValue, Double) * CType(ADUValue, Double)) * CType(ValueCount, Double)                  'ADUValue^2 * count
                MeanSum += WeightCount
                MeanPow2Sum += WeightPow2
                If ValueCount > RetVal.Modus.Value Then RetVal.Modus = New KeyValuePair(Of Int64, UInt64)(ADUValue, Histogram(ADUValue))    'modus (most "used" ADU value)
                If SamplesSeen >= RetVal.Samples / 2 And RetVal.Median = Int64.MinValue Then RetVal.Median = ADUValue                       'median value (set once)
                Dim PctIdx As Integer = CInt(Math.Round(100 * (SamplesSeen / RetVal.Samples)))                                              'percentile index (0...100)
                If PctIdx >= 25 And PctIdx <= 75 Then RetVal.ADUValues2575 += 1                                                             'number of different ADU counts in percentile range 25..75
            Next ADUValue

            '--------------------------------------------------------------------------------------------------------------------------------------------
            'Percentile calculation
            '--------------------------------------------------------------------------------------------------------------------------------------------

            '1.) Init percentile - percentiles are writen in each bin as an incremental processing fails in fast-changing histograms
            Dim PCTInvalid As Long = Long.MinValue
            For Pct As Integer = 0 To 100
                RetVal.Percentile.Add(Pct, PCTInvalid)
            Next Pct

            'Move over the histogram for percentile and values in 25-75pct range
            Dim NextPctIdx As Integer = 1
            SamplesSeen = 0
            For Each ADUValue As ADUFixed In AllADUValues
                Dim NextPctLimit As UInt64 = CType(NextPctIdx * (RetVal.Samples / 100), UInt64)                                             'calculate in every round not required but makes it easier to understand
                SamplesSeen += Histogram(ADUValue)
                If SamplesSeen >= NextPctLimit Then
                    RetVal.Percentile(NextPctIdx) = ADUValue
                    NextPctIdx += 1
                End If
            Next ADUValue

            'Set percentiles in bin which to not have a valid entry
            Dim LastValidPct As Long = RetVal.Min.Key
            For Pct As Integer = 0 To 100
                If RetVal.Percentile(Pct) = PCTInvalid Then
                    RetVal.Percentile(Pct) = LastValidPct
                Else
                    LastValidPct = RetVal.Percentile(Pct)
                End If
            Next Pct

            'Calculate the quantizer histogram
            RetVal.HistXDist = GetQuantizationHisto(Histogram)

            'Calculate final outputs
            RetVal.Mean = MeanSum / RetVal.Samples
            RetVal.MeanPow2 = MeanPow2Sum / RetVal.Samples
            RetVal.StdDev = Math.Sqrt(RetVal.MeanPow2 - (RetVal.Mean * RetVal.Mean))
            Return RetVal

        End Function

        '''<summary>Calculate the statistic data from the passed histogram data.</summary>
        '''<param name="Histogram">Calculated histogram data.</param>
        Private Shared Function CalcStatisticFromHistogram(ByRef Histogram As Dictionary(Of Single, ADUCount)) As sSingleChannelStatistics_Float32

            If IsNothing(Histogram) = True Then Return Nothing

            Dim RetVal As sSingleChannelStatistics_Float32 = sSingleChannelStatistics_Float32.Init()
            Dim AllADUValues As List(Of Single) = Histogram.KeyList
            AllADUValues.Sort()

            'Count number of samples
            For Each PixelValue As Single In Histogram.Keys
                RetVal.Samples += Histogram(PixelValue)
            Next PixelValue

            'Store number of different sample values
            RetVal.DifferentADUValues = Histogram.Count

            'Init statistics calculation
            Dim SumSampleCount As UInt64 = 0
            Dim MeanSum As Double = 0
            Dim MeanPow2Sum As Double = 0
            RetVal.Min = New KeyValuePair(Of Single, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.Max = New KeyValuePair(Of Single, UInt64)(AllADUValues(AllADUValues.Count - 1), Histogram(AllADUValues(AllADUValues.Count - 1)))
            RetVal.Modus = New KeyValuePair(Of Single, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.HistXDist = New Dictionary(Of Single, UInt64)

            'Init percentile - percentiles are writen in each bin as an incremental processing fails in fast-changing histograms
            Dim PCTInvalid As Long = Long.MinValue
            For Pct As Integer = 0 To 100
                RetVal.Percentile.Add(Pct, PCTInvalid)
            Next Pct

            'Move over the histogram for percentile and values in 25-75pct range
            RetVal.ADUValues2575 = 0
            For Each ADUValue As Single In AllADUValues
                Dim ValueCount As ADUCount = Histogram(ADUValue)
                SumSampleCount += ValueCount
                Dim WeightCount As Double = ADUValue * ADUValue
                Dim WeightPow2 As Double = (ADUValue * ADUValue) * CType(ValueCount, Double)
                MeanSum += WeightCount
                MeanPow2Sum += WeightPow2
                If ValueCount > RetVal.Modus.Value Then RetVal.Modus = New KeyValuePair(Of Single, UInt64)(ADUValue, Histogram(ADUValue))
                If SumSampleCount >= RetVal.Samples / 2 And RetVal.Median = Int64.MinValue Then RetVal.Median = ADUValue
                Dim PctIdx As Integer = CInt(100 * (SumSampleCount / RetVal.Samples))
                If RetVal.Percentile(PctIdx) = PCTInvalid Then RetVal.Percentile(PctIdx) = ADUValue
                If PctIdx >= 25 And PctIdx <= 75 Then RetVal.ADUValues2575 += 1
            Next ADUValue
            RetVal.HistXDist = GetQuantizationHisto(Histogram)

            'Set percentiles in bin which to not have a valid entry
            Dim LastValidPct As Single = RetVal.Min.Key
            For Pct As Integer = 0 To 100
                If RetVal.Percentile(Pct) = PCTInvalid Then
                    RetVal.Percentile(Pct) = LastValidPct
                Else
                    LastValidPct = RetVal.Percentile(Pct)
                End If
            Next Pct

            'Calculate final outputs
            RetVal.Mean = MeanSum / RetVal.Samples
            RetVal.MeanPow2 = MeanPow2Sum / RetVal.Samples
            RetVal.StdDev = Math.Sqrt(RetVal.MeanPow2 - (RetVal.Mean * RetVal.Mean))
            Return RetVal

        End Function

        '''<summary>Get the histogram for all quanization level differences found.</summary>
        '''<param name="Histo">Histogram data with ADU value and number of pixel with this ADU value.</param>
        Public Shared Function GetQuantizationHisto(ByRef Histo As Dictionary(Of ADUFixed, ADUCount)) As Dictionary(Of Long, UInt64)
            Dim RetVal As New Dictionary(Of Long, UInt64)
            Dim LastHistX As ADUFixed = Int64.MaxValue
            For Each HistoX As ADUFixed In Histo.KeyList
                If LastHistX <> ADUFixed.MaxValue Then
                    Dim Distance As Long = CType(HistoX - LastHistX, Long)
                    RetVal.AddTo(Distance, UInt64_1)
                End If
                LastHistX = HistoX
            Next HistoX
            Return RetVal.SortDictionary
        End Function

        '''<summary>Get the histogram for all quanization level differences found.</summary>
        '''<param name="Histo">Histogram data with ADU value and number of pixel with this ADU value.</param>
        Public Shared Function GetQuantizationHisto(ByRef Histo As Dictionary(Of Single, ADUCount)) As Dictionary(Of Single, UInt64)
            Dim RetVal As New Dictionary(Of Single, UInt64)
            Dim LastHistX As Single = Single.NaN
            For Each HistoX As Single In Histo.KeyList
                If Single.IsNaN(LastHistX) = False Then
                    Dim Distance As Single = HistoX - LastHistX
                    RetVal.AddTo(Distance, 1)
                End If
                LastHistX = HistoX
            Next HistoX
            Return RetVal.SortDictionary
        End Function

        '''<summary>Combine all bayer statistics to a monochromatic statistic of all pixel of the image.</summary>
        Public Shared Function CombineBayerToMonoStatistics(Of T)(ByRef BayerHistData(,) As Dictionary(Of T, ADUCount)) As Dictionary(Of T, ADUCount)
            Dim RetVal As New Dictionary(Of T, ADUCount)
            For Idx1 As Integer = 0 To BayerHistData.GetUpperBound(0)
                For Idx2 As Integer = 0 To BayerHistData.GetUpperBound(1)
                    If IsNothing(BayerHistData(Idx1, Idx2)) = False Then
                        For Each KeyIdx As T In BayerHistData(Idx1, Idx2).Keys
                            RetVal.AddTo(KeyIdx, BayerHistData(Idx1, Idx2)(KeyIdx))
                        Next KeyIdx
                    End If
                Next Idx2
            Next Idx1
            Return RetVal.SortDictionary
        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="XEntries">Number of different X entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<param name="YEntries">Number of different Y entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Function BayerStatistics_Int(ByVal NAXIS3 As Integer) As Dictionary(Of ADUFixed, ADUCount)(,)

            'Count all values
            Dim RetVal(1, 1) As Dictionary(Of ADUFixed, ADUCount)

            'Data are UInt16
            If IsNothing(DataProcessor_UInt16) = False Then
                If IsNothing(DataProcessor_UInt16.ImageData) = False Then
                    If DataProcessor_UInt16.ImageData(NAXIS3).Length > 0 Then
                        Dim Results(,) As cStatMultiThread.cStatObjFixed = Nothing
                        DataProcessor_UInt16.Calculate(NAXIS3, Results)
                        For Idx1 As Integer = 0 To 1
                            For Idx2 As Integer = 0 To 1
                                RetVal(Idx1, Idx2) = Results(Idx1, Idx2).HistDataBayer
                            Next Idx2
                        Next Idx1
                    End If
                End If
            End If

            'Data are Int32
            If IsNothing(DataProcessor_Int32) = False Then
                If IsNothing(DataProcessor_Int32.ImageData) = False Then
                    If DataProcessor_Int32.ImageData.Length > 0 Then
                        Dim Results(,) As cStatMultiThread.cStatObjFixed = Nothing
                        DataProcessor_Int32.Calculate(Results)
                        For Idx1 As Integer = 0 To 1
                            For Idx2 As Integer = 0 To 1
                                RetVal(Idx1, Idx2) = Results(Idx1, Idx2).HistDataBayer
                            Next Idx2
                        Next Idx1
                    End If
                End If
            End If

            'Data are UInt32
            If IsNothing(DataProcessor_UInt32) = False Then
                If IsNothing(DataProcessor_UInt32.ImageData) = False Then
                    If DataProcessor_UInt32.ImageData(NAXIS3).Length > 0 Then
                        Dim Results(,) As cStatMultiThread.cStatObjFixed = Nothing
                        DataProcessor_UInt32.Calculate(NAXIS3, Results)
                        For Idx1 As Integer = 0 To 1
                            For Idx2 As Integer = 0 To 1
                                RetVal(Idx1, Idx2) = Results(Idx1, Idx2).HistDataBayer
                            Next Idx2
                        Next Idx1
                    End If
                End If
            End If

            Return RetVal

        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="XEntries">Number of different X entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<param name="YEntries">Number of different Y entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Function BayerStatistics_Float32() As Dictionary(Of Single, ADUCount)(,)

            'Count all values
            Dim RetVal(1, 1) As Dictionary(Of Single, UInt64)

            'Data are Float32
            If IsNothing(DataProcessor_Float32) = False Then
                If IsNothing(DataProcessor_Float32.ImageData) = False Then
                    If DataProcessor_Float32.ImageData.Length > 0 Then
                        Dim Results(,) As cStatMultiThread_Float32.cStateObj = Nothing
                        DataProcessor_Float32.Calculate(Results)
                        For Idx1 As Integer = 0 To 1
                            For Idx2 As Integer = 0 To 1
                                RetVal(Idx1, Idx2) = Results(Idx1, Idx2).HistDataBayer
                            Next Idx2
                        Next Idx1
                    End If
                End If
            End If

            Return RetVal

        End Function

    End Class

End Namespace