Option Explicit On
Option Strict On

'This class is the container for all astronomical statistics functions
'All other modules with statistics should not be used any more and put their code here
'Old modules:
' - C:\GIT\src\atomics\cStatistics.vb
Namespace AstroNET

    Public Class Statistics

        '''<summary>Instance of Intel IPP library call.</summary>
        Private IntelIPP As cIntelIPP = Nothing

        '''<summary>Image data in UInt16 mode.</summary>
        Public DataProcessor_UInt16 As New cStatMultiThread_UInt16

        '''<summary>Image data in Int32 mode.</summary>
        Public DataProcessor_Int32 As New cStatMultiThread_Int32

        '''<summary>Image data in UInt32 mode.</summary>
        Public DataProcessor_UInt32 As New cStatMultiThread_UInt32

        Private Const UInt64_1 As UInt64 = CType(1, UInt64)

        Public ReadOnly Property DataProcessorUsed() As String
            Get
                If DataProcessor_UInt16.ImageData.LongLength > 0 Then Return GetType(UInt16).Name
                If DataProcessor_UInt32.ImageData.LongLength > 0 Then Return GetType(UInt32).Name
                If DataProcessor_Int32.ImageData.LongLength > 0 Then Return GetType(Int32).Name
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
            '''<summary>Full-resolution histogram data - bayer data.</summary>
            Public BayerHistograms(,) As Collections.Generic.Dictionary(Of Int64, UInt64)
            '''<summary>Full-resolution histogram data - mono data, sorted.</summary>
            Public MonochromHistogram As Collections.Generic.Dictionary(Of Int64, UInt64)
            '''<summary>Statistics for each channel.</summary>
            Public BayerStatistics(,) As sSingleChannelStatistics
            '''<summary>Statistics for each channel.</summary>
            Public MonoStatistics As sSingleChannelStatistics
            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport() As Collections.Generic.List(Of String)
                Return StatisticsReport(String.Empty)
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport(ByVal Indent As String) As Collections.Generic.List(Of String)
                Dim RetVal As New Collections.Generic.List(Of String)
                RetVal.Add(Indent & "Property".PadRight(sSingleChannelStatistics.ReportHeaderLength) & ": " & "Mono".PadRight(sSingleChannelStatistics.ReportValueLength) & "|")
                For Each Entry As String In MonoStatistics.StatisticsReport
                    RetVal.Add(Indent & "  " & Entry & "|")
                Next Entry
                For Idx1 As Integer = 0 To BayerStatistics.GetUpperBound(0)
                    For Idx2 As Integer = 0 To BayerStatistics.GetUpperBound(1)
                        RetVal(0) &= ("Bay[" & Idx1.ValRegIndep & ":" & Idx2.ValRegIndep & "]").PadRight(sSingleChannelStatistics.ReportValueLength) & "|"
                        Dim LineIdx As Integer = 1
                        For Each Entry As String In BayerStatistics(Idx1, Idx2).StatisticsReport
                            RetVal(LineIdx) &= Entry.Substring(sSingleChannelStatistics.ReportHeaderLength) & "|"
                            LineIdx += 1
                        Next Entry
                    Next Idx2
                Next Idx1
                Return RetVal
            End Function
        End Structure

        '''<summary>Statistic information of one channel (RGB or total).</summary>
        '''<remarks>The maximum word with is taken as pixel values to cover all fixed-point formats ...</remarks>
        Public Structure sSingleChannelStatistics
            '''<summary>Number of characters in the header of the report.</summary>
            Public Shared ReadOnly Property ReportHeaderLength As Integer = 20
            '''<summary>Number of characters in the value of the report.</summary>
            Public Shared ReadOnly Property ReportValueLength As Integer = 16
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As UInt64
            '''<summary>Maximum value occured (value and number of pixel that have this value).</summary>
            Public Max As Collections.Generic.KeyValuePair(Of Int64, UInt64)
            '''<summary>Minimum value occured (value and number of pixel that have this value).</summary>
            Public Min As Collections.Generic.KeyValuePair(Of Int64, UInt64)
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
            Public HistXDist As Collections.Generic.Dictionary(Of Long, UInt64)
            '''<summary>Percentile.</summary>
            Public Percentile As Collections.Generic.Dictionary(Of Integer, Int64)
            '''<summary>Pixel value that is present the most often.</summary>
            Public Modus As Collections.Generic.KeyValuePair(Of Int64, UInt64)
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public ReadOnly Property Variance As Double
                Get
                    Return StdDev ^ 2
                End Get
            End Property
            '''<summary>Init all inner variables.</summary>
            Public Shared Function InitForShort() As sSingleChannelStatistics
                Dim RetVal As New sSingleChannelStatistics
                RetVal.Samples = 0
                RetVal.Max = Nothing
                RetVal.Min = Nothing
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentADUValues = 0
                RetVal.HistXDist = New Collections.Generic.Dictionary(Of Long, UInt64)
                RetVal.Median = Int64.MinValue
                RetVal.Percentile = New Collections.Generic.Dictionary(Of Integer, Int64)
                RetVal.Modus = Nothing
                Return RetVal
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            '''<param name="DispHeader">TRUE to display the header, FALSE else.</param>
            Public Function StatisticsReport() As Collections.Generic.List(Of String)
                Dim NotPresent As String = New String("-"c, ReportValueLength)
                Dim RetVal As New Collections.Generic.List(Of String)
                Dim HistXDist_keys As Collections.Generic.List(Of Long) = HistXDist.KeyList
                RetVal.Add("Total pixel       : " & Samples.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Total pixel       : " & ((Samples / 1000000).ValRegIndep("0.0") & "M").PadLeft(ReportValueLength))
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
        Public Function ImageStatistics() As sStatistics

            Dim RetVal As New sStatistics

            RetVal.BayerHistograms = BayerStatistics()      'Calculate a 2x2 bayer statistics (also for mono data as thread-based will speed up ...)
            CalculateAllFromBayerStatistics(RetVal)         'Add all other data (mono histo and statistics)

            'Return results
            Return RetVal

        End Function

        '''<summary>Combine 2 SingleChannelStatistics elements (e.g. to calculate the aggregated statistic for multi-frame capture).</summary>
        Public Shared Function CombineStatistics(ByVal StatA As sStatistics, ByVal CombinedStatistics As sStatistics) As sStatistics
            Dim RetVal As New sStatistics
            '1.) Combine to 2 histograms
            ReDim RetVal.BayerHistograms(StatA.BayerHistograms.GetUpperBound(0), StatA.BayerHistograms.GetUpperBound(1))
            For BayIdx1 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(0)
                For BayIdx2 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(1)
                    RetVal.BayerHistograms(BayIdx1, BayIdx2) = New Collections.Generic.Dictionary(Of Int64, UInt64)
                    'Init return bayer histogram with StatA data
                    For Each PixelValue As UInt16 In StatA.BayerHistograms(BayIdx1, BayIdx2).Keys
                        RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, StatA.BayerHistograms(BayIdx1, BayIdx2)(PixelValue))
                    Next PixelValue
                    'Combine with StatB data
                    If IsNothing(CombinedStatistics.BayerHistograms) = False Then
                        For Each PixelValue As UInt16 In CombinedStatistics.BayerHistograms(BayIdx1, BayIdx2).Keys
                            Dim HistoCount As UInt64 = CombinedStatistics.BayerHistograms(BayIdx1, BayIdx2)(PixelValue)
                            If RetVal.BayerHistograms(BayIdx1, BayIdx2).ContainsKey(PixelValue) = False Then
                                RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, HistoCount)
                            Else
                                RetVal.BayerHistograms(BayIdx1, BayIdx2)(PixelValue) += HistoCount
                            End If
                        Next PixelValue
                    End If
                    RetVal.BayerHistograms(BayIdx1, BayIdx2) = RetVal.BayerHistograms(BayIdx1, BayIdx2).SortDictionary
                Next BayIdx2
            Next BayIdx1
            CalculateAllFromBayerStatistics(RetVal)
            Return RetVal
        End Function

        '''<summary>Calculate all statistic data (mono histo and statistics) from the passed bayer statistics.</summary>
        Private Shared Sub CalculateAllFromBayerStatistics(ByRef RetVal As sStatistics)
            'Calculate a monochromatic statistics from the bayer histograms
            RetVal.MonochromHistogram = CombineBayerToMonoStatistics(RetVal.BayerHistograms)
            'Calculate the bayer channel statistics from the bayer histogram
            ReDim RetVal.BayerStatistics(RetVal.BayerHistograms.GetUpperBound(0), RetVal.BayerHistograms.GetUpperBound(1))
            For Idx1 As Integer = 0 To RetVal.BayerHistograms.GetUpperBound(0)
                For Idx2 As Integer = 0 To RetVal.BayerHistograms.GetUpperBound(1)
                    RetVal.BayerStatistics(Idx1, Idx2) = CalcStatisticFromHistogram(RetVal.BayerHistograms(Idx1, Idx2))
                Next Idx2
            Next Idx1
            'Calculate the total statistics
            RetVal.MonoStatistics = CalcStatisticFromHistogram(RetVal.MonochromHistogram)
        End Sub

        '''<summary>Calculate the statistic data from the passed histogram data.</summary>
        '''<param name="Histogram">Calculated histogram data.</param>
        Private Shared Function CalcStatisticFromHistogram(ByRef Histogram As Collections.Generic.Dictionary(Of Int64, UInt64)) As sSingleChannelStatistics

            Dim RetVal As sSingleChannelStatistics = sSingleChannelStatistics.InitForShort()
            Dim AllADUValues As Collections.Generic.List(Of Int64) = Histogram.KeyList
            AllADUValues.Sort()

            'Count number of samples
            For Each PixelValue As Int64 In Histogram.Keys
                RetVal.Samples += Histogram(PixelValue)
            Next PixelValue

            'Store number of different sample values
            RetVal.DifferentADUValues = Histogram.Count

            'Init statistics calculation
            Dim SumSampleCount As UInt64 = 0
            Dim MeanSum As Double = 0
            Dim MeanPow2Sum As System.Double = 0
            RetVal.Min = New Collections.Generic.KeyValuePair(Of Int64, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.Max = New Collections.Generic.KeyValuePair(Of Int64, UInt64)(AllADUValues(AllADUValues.Count - 1), Histogram(AllADUValues(AllADUValues.Count - 1)))
            RetVal.Modus = New Collections.Generic.KeyValuePair(Of Int64, UInt64)(AllADUValues(0), Histogram(AllADUValues(0)))
            RetVal.HistXDist = New Collections.Generic.Dictionary(Of Long, UInt64)

            'Init percentile - percentiles are writen in each bin as an incremental processing fails in fast-changing histograms
            Dim PCTInvalid As Long = Long.MinValue
            For Pct As Integer = 0 To 100
                RetVal.Percentile.Add(Pct, PCTInvalid)
            Next Pct

            'Move over the histogram for percentile and values in 25-75pct range
            RetVal.ADUValues2575 = 0
            For Each ADUValue As Int64 In AllADUValues
                Dim ValueCount As UInt64 = Histogram(ADUValue)
                SumSampleCount += ValueCount
                Dim WeightCount As Double = (CType(ADUValue, Double) * CType(ValueCount, Double))
                Dim WeightPow2 As Double = (CType(ADUValue, Double) * CType(ADUValue, Double)) * CType(ValueCount, Double)
                MeanSum += WeightCount
                MeanPow2Sum += WeightPow2
                If ValueCount > RetVal.Modus.Value Then RetVal.Modus = New Collections.Generic.KeyValuePair(Of Int64, UInt64)(ADUValue, Histogram(ADUValue))
                If SumSampleCount >= RetVal.Samples / 2 And RetVal.Median = Int64.MinValue Then RetVal.Median = ADUValue
                Dim PctIdx As Integer = CInt(100 * (SumSampleCount / RetVal.Samples))
                If RetVal.Percentile(PctIdx) = PCTInvalid Then RetVal.Percentile(PctIdx) = ADUValue
                If PctIdx >= 25 And PctIdx <= 75 Then RetVal.ADUValues2575 += 1
            Next ADUValue
            RetVal.HistXDist = GetQuantizationHisto(Histogram)

            'Set percentiles in bin which to not have a valid entry
            Dim LastValidPct As Long = RetVal.Min.Key
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
        Public Shared Function GetQuantizationHisto(ByRef Histo As Collections.Generic.Dictionary(Of Long, UInt64)) As Collections.Generic.Dictionary(Of Long, UInt64)
            Dim RetVal As New Collections.Generic.Dictionary(Of Long, UInt64)
            Dim LastHistX As Int64 = Int64.MaxValue
            For Each HistoX As Int64 In Histo.KeyList
                If LastHistX <> Int64.MaxValue Then
                    Dim Distance As UInteger = CUInt(HistoX - LastHistX)
                    If RetVal.ContainsKey(Distance) = False Then
                        RetVal.Add(Distance, 1)
                    Else
                        RetVal(Distance) = RetVal(Distance) + UInt64_1
                    End If
                End If
                LastHistX = HistoX
            Next HistoX
            Return RetVal.SortDictionary
        End Function

        '''<summary>Get the histogram for all quanization level differences found.</summary>
        '''<param name="Histo">Histogram data with ADU value and number of pixel with this ADU value.</param>
        Public Shared Function GetQuantizationHisto(ByRef Histo As Collections.Generic.Dictionary(Of Single, UInt32)) As Collections.Generic.Dictionary(Of Single, UInt32)
            Dim RetVal As New Collections.Generic.Dictionary(Of Single, UInt32)
            Dim LastHistX As Single = Single.NaN
            For Each HistoX As Single In Histo.KeyList
                If Single.IsNaN(LastHistX) = False Then
                    Dim Distance As Single = HistoX - LastHistX
                    If RetVal.ContainsKey(Distance) = False Then
                        RetVal.Add(Distance, 1)
                    Else
                        RetVal(Distance) = CUInt(RetVal(Distance) + 1)
                    End If
                End If
                LastHistX = HistoX
            Next HistoX
            Return RetVal.SortDictionary
        End Function

        '''<summary>Combine all bayer statistics to a monochromatic statistic of all pixel of the image.</summary>
        Public Shared Function CombineBayerToMonoStatistics(Of T)(ByRef BayerHistData(,) As Collections.Generic.Dictionary(Of T, UInt64)) As Collections.Generic.Dictionary(Of T, UInt64)
            Dim RetVal As New Collections.Generic.Dictionary(Of T, UInt64)
            For Idx1 As Integer = 0 To BayerHistData.GetUpperBound(0)
                For Idx2 As Integer = 0 To BayerHistData.GetUpperBound(1)
                    If IsNothing(BayerHistData(Idx1, Idx2)) = False Then
                        For Each KeyIdx As T In BayerHistData(Idx1, Idx2).Keys
                            If RetVal.ContainsKey(KeyIdx) = False Then
                                RetVal.Add(KeyIdx, BayerHistData(Idx1, Idx2)(KeyIdx))
                            Else
                                RetVal(KeyIdx) += BayerHistData(Idx1, Idx2)(KeyIdx)
                            End If
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
        Public Function BayerStatistics() As Collections.Generic.Dictionary(Of Int64, UInt64)(,)

            'Count all values
            Dim RetVal(1, 1) As Collections.Generic.Dictionary(Of Int64, UInt64)

            'Data are UInt16
            If IsNothing(DataProcessor_UInt16) = False Then
                If IsNothing(DataProcessor_UInt16.ImageData) = False Then
                    If DataProcessor_UInt16.ImageData.Length > 0 Then
                        Dim Results(,) As cStatMultiThread_UInt16.cStateObj = Nothing
                        DataProcessor_UInt16.Calculate(Results)
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
                        Dim Results(,) As cStatMultiThread_Int32.cStateObj = Nothing
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
                    If DataProcessor_UInt32.ImageData.Length > 0 Then
                        Dim Results(,) As cStatMultiThread_UInt32.cStateObj = Nothing
                        DataProcessor_UInt32.Calculate(Results)
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