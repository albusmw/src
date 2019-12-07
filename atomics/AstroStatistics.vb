Option Explicit On
Option Strict On

'This class is the container for all astronomical statistics functions
'All other modules with statistics should not be used any more and put their code here
'Old modules:
' - C:\GIT\src\atomics\cStatistics.vb
Namespace AstroNET

    Public Class Statistics

        '''<summary>Path to ipps.dll and ippvm.dll - if not set IPP will not be used.</summary>
        Public Shared Property IPPPath As String = String.Empty

        '''<summary>Instance of Intel IPP library call.</summary>
        Private IntelIPP As cIntelIPP = Nothing

        '''<summary>Image data in UInt16 mode.</summary>
        Public DataProcessor_UInt16 As New cStatMultiThread_UInt16

        '''<summary>Image data in UInt32 mode.</summary>
        Public DataProcessor_UInt32 As New cStatMultiThread_UInt32

        Private Const OneUInt32 As UInt32 = CType(1, UInt32)

        Public Sub New()
            IntelIPP = New cIntelIPP(System.IO.Path.Combine(IPPPath, "ipps.dll"), System.IO.Path.Combine(IPPPath, "ippvm.dll"), System.IO.Path.Combine(IPPPath, "ippi.dll"))
        End Sub

        '''<summary>Total statistics - available as per-channel bayer statistics and as combined statistics.</summary>
        Public Structure sStatistics
            '''<summary>Full-resolution histogram data - bayer data.</summary>
            Public BayerHistograms(,) As Dictionary(Of UInt32, UInt32)
            '''<summary>Full-resolution histogram data - mono data.</summary>
            Public MonochromHistogram As Dictionary(Of UInt32, UInt32)
            '''<summary>Statistics for each channel.</summary>
            Public BayerStatistics(,) As sSingleChannelStatistics
            '''<summary>Statistics for each channel.</summary>
            Public MonoStatistics As sSingleChannelStatistics
            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport() As List(Of String)
                Return StatisticsReport(String.Empty)
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport(ByVal Indent As String) As List(Of String)
                Dim RetVal As New List(Of String)
                RetVal.Add(Indent & "Total (mono) statistics:")
                For Each Entry As String In MonoStatistics.StatisticsReport
                    RetVal.Add(Indent & "  " & Entry)
                Next Entry
                For Idx1 As Integer = 0 To BayerStatistics.GetUpperBound(0)
                    For Idx2 As Integer = 0 To BayerStatistics.GetUpperBound(1)
                        RetVal.Add(Indent & "Bayer channel [" & Idx1.ToString.Trim & ":" & Idx2.ToString.Trim & "] statistics:")
                        For Each Entry As String In BayerStatistics(Idx1, Idx2).StatisticsReport
                            RetVal.Add(Indent & "  " & Entry)
                        Next Entry
                    Next Idx2
                Next Idx1
                Return RetVal
            End Function
        End Structure

        '''<summary>Statistic information of one channel (RGB or total).</summary>
        Public Structure sSingleChannelStatistics
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As Long
            '''<summary>Maximum value occured.</summary>
            Public Max As UInt32
            '''<summary>Minimum value occured.</summary>
            Public Min As UInt32
            '''<summary>Value where half of the samples are below and half are above.</summary>
            Public Median As UInt32
            '''<summary>Arithmetic mean value.</summary>
            Public Mean As Double
            '''<summary>Mean value of squared values.</summary>
            Public MeanPow2 As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public StdDev As Double
            '''<summary>Number of different values in the data.</summary>
            Public DifferentValueCount As Integer
            '''<summary>Percentile.</summary>
            Public Percentile As Dictionary(Of Byte, UInt32)
            '''<summary>Init all inner variables.</summary>
            Public Shared Function InitForShort() As sSingleChannelStatistics
                Dim RetVal As New sSingleChannelStatistics
                RetVal.Samples = 0
                RetVal.Max = UInt32.MinValue
                RetVal.Min = UInt32.MaxValue
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentValueCount = 0
                RetVal.Median = UInt32.MinValue
                RetVal.Percentile = New Dictionary(Of Byte, UInt32)
                Return RetVal
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            Public Function StatisticsReport() As List(Of String)
                Dim RetVal As New List(Of String)
                RetVal.Add("Total pixel     : " & Samples.ToString.Trim.PadLeft(9) & " (" & Format(Samples / 1000000, "0.0").ToString.Trim & "M)")
                RetVal.Add("Different values: " & DifferentValueCount.ToString.Trim.PadLeft(9))
                RetVal.Add("Min value       : " & Min.ToString.Trim.PadLeft(9))
                RetVal.Add("Max value       : " & Max.ToString.Trim.PadLeft(9))
                RetVal.Add("Median value    : " & Median.ToString.Trim.PadLeft(9))
                RetVal.Add("Mean value      : " & Format(Mean, "0.00").ToString.Trim.PadLeft(9))
                Return RetVal
            End Function
        End Structure

        '''<summary>Calculate the image statistics of the passed image data.</summary>
        Public Function ImageStatistics() As sStatistics

            Dim RetVal As New sStatistics

            Dim Stopper As New Stopwatch : Stopper.Reset() : Stopper.Start()

            'Calculate a 2x2 bayer statistics (also for mono data - if thread-based this may even speed up ...)
            RetVal.BayerHistograms = BayerStatistics()
            Stopper.Stop()
            Dim T1 As Long = Stopper.ElapsedMilliseconds

            'Add all other data (mono histo and statistics)
            CalculateAllFromBayerStatistics(RetVal)

            'Return results
            Return RetVal

        End Function

        '''<summary>Combine 2 SingleChannelStatistics elements (e.g. to calculate the aggregated statistic for multi-frame capture).</summary>
        Public Function CombineStatistics(ByVal StatA As sStatistics, ByVal StatB As sStatistics) As sStatistics
            Dim RetVal As New sStatistics
            '1.) Combine to 2 histograms
            ReDim RetVal.BayerHistograms(StatA.BayerHistograms.GetUpperBound(0), StatA.BayerHistograms.GetUpperBound(1))
            For BayIdx1 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(0)
                For BayIdx2 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(1)
                    RetVal.BayerHistograms(BayIdx1, BayIdx2) = New Dictionary(Of UInt32, UInteger)
                    'Init return bayer histogram with StatA data
                    For Each PixelValue As UInt16 In StatA.BayerHistograms(BayIdx1, BayIdx2).Keys
                        RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, StatA.BayerHistograms(BayIdx1, BayIdx2)(PixelValue))
                    Next PixelValue
                    'Combine with StatB data
                    If IsNothing(StatB.BayerHistograms) = False Then
                        For Each PixelValue As UInt16 In StatB.BayerHistograms(BayIdx1, BayIdx2).Keys
                            Dim HistoCount As UInteger = StatB.BayerHistograms(BayIdx1, BayIdx2)(PixelValue)
                            If RetVal.BayerHistograms(BayIdx1, BayIdx2).ContainsKey(PixelValue) = False Then
                                RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, HistoCount)
                            Else
                                RetVal.BayerHistograms(BayIdx1, BayIdx2)(PixelValue) += HistoCount
                            End If
                        Next PixelValue
                    End If
                    RetVal.BayerHistograms(BayIdx1, BayIdx2) = cGenerics.SortDictionary(RetVal.BayerHistograms(BayIdx1, BayIdx2))
                Next BayIdx2
            Next BayIdx1
            CalculateAllFromBayerStatistics(RetVal)
            Return RetVal
        End Function

        '''<summary>Calculate all statistic data (mono histo and statistics) from the passed bayer statistics.</summary>
        Private Sub CalculateAllFromBayerStatistics(ByRef RetVal As sStatistics)
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
        Private Function CalcStatisticFromHistogram(ByRef Histogram As Dictionary(Of UInt32, UInt32)) As sSingleChannelStatistics
            Dim RetVal As sSingleChannelStatistics = sSingleChannelStatistics.InitForShort()
            Dim SamplesProcessed As UInt32 = 0
            'Count number of samples
            For Each PixelValue As UInt32 In Histogram.Keys
                RetVal.Samples += Histogram(PixelValue)
            Next PixelValue
            'Store number of different sample values
            RetVal.DifferentValueCount = Histogram.Count
            'Calculate statistics
            Dim SumSampleCount As Long = 0
            Dim MeanSum As UInt64 = 0
            Dim MeanPow2Sum As System.Double = 0
            Dim Lim_Pct5 As Long = CLng(RetVal.Samples * 0.05)
            Dim Lim_Pct25 As Long = CLng(RetVal.Samples * 0.25)
            Dim Lim_Pct50 As Long = CLng(RetVal.Samples * 0.5)
            Dim Lim_Pct75 As Long = CLng(RetVal.Samples * 0.75)
            Dim Lim_Pct95 As Long = CLng(RetVal.Samples * 0.95)
            For Each PixelValue As UInt32 In Histogram.Keys
                Dim HistCount As UInteger = Histogram(PixelValue)
                SumSampleCount += HistCount
                Dim WeightCount As UInt64 = (CType(PixelValue, UInt64) * CType(HistCount, UInt64))
                Dim WeightPow2 As UInt64 = (CType(PixelValue, UInt64) * CType(PixelValue, UInt64)) * CType(HistCount, UInt64)
                SamplesProcessed += HistCount
                MeanSum += WeightCount
                MeanPow2Sum += WeightPow2
                If PixelValue > RetVal.Max Then RetVal.Max = PixelValue
                If PixelValue < RetVal.Min Then RetVal.Min = PixelValue
                If SamplesProcessed >= RetVal.Samples \ 2 And RetVal.Median = UInt16.MinValue Then RetVal.Median = PixelValue
                If SumSampleCount >= Lim_Pct5 And RetVal.Percentile.ContainsKey(5) = False Then RetVal.Percentile.Add(5, PixelValue)
                If SumSampleCount >= Lim_Pct25 And RetVal.Percentile.ContainsKey(25) = False Then RetVal.Percentile.Add(25, PixelValue)
                If SumSampleCount >= Lim_Pct50 And RetVal.Percentile.ContainsKey(50) = False Then RetVal.Percentile.Add(50, PixelValue)
                If SumSampleCount >= Lim_Pct75 And RetVal.Percentile.ContainsKey(75) = False Then RetVal.Percentile.Add(75, PixelValue)
                If SumSampleCount >= Lim_Pct95 And RetVal.Percentile.ContainsKey(95) = False Then RetVal.Percentile.Add(95, PixelValue)
            Next PixelValue
            RetVal.StdDev = Math.Sqrt(((RetVal.MeanPow2) - ((RetVal.Mean * RetVal.Mean) / RetVal.Samples)) / (RetVal.Samples - 1))
            RetVal.Mean = MeanSum / RetVal.Samples
            RetVal.MeanPow2 = MeanPow2Sum / RetVal.Samples
            Return RetVal
        End Function

        '''<summary>Combine all bayer statistics to a monochromatic statistic of all pixel of the image.</summary>
        Public Function CombineBayerToMonoStatistics(Of T)(ByRef BayerHistData(,) As Dictionary(Of T, UInt32)) As Dictionary(Of T, UInt32)
            Dim RetVal As New Dictionary(Of T, UInt32)
            For Idx1 As Integer = 0 To BayerHistData.GetUpperBound(0)
                For Idx2 As Integer = 0 To BayerHistData.GetUpperBound(1)
                    For Each KeyIdx As T In BayerHistData(Idx1, Idx2).Keys
                        If RetVal.ContainsKey(KeyIdx) = False Then
                            RetVal.Add(KeyIdx, BayerHistData(Idx1, Idx2)(KeyIdx))
                        Else
                            RetVal(KeyIdx) += BayerHistData(Idx1, Idx2)(KeyIdx)
                        End If
                    Next KeyIdx
                Next Idx2
            Next Idx1
            Return cGenerics.SortDictionary(RetVal)
        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="XEntries">Number of different X entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<param name="YEntries">Number of different Y entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Function BayerStatistics() As Dictionary(Of UInt32, UInt32)(,)

            'Count all values
            Dim RetVal(1, 1) As Dictionary(Of UInt32, UInt32)

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

        ''' <summary>Calculate the intensity over the distance from the center of the image.</summary>
        ''' <param name="FITSSumImage">Image to run calculation on.</param>
        ''' <param name="Steps">Number of X axis steps to group - 0 for full resolution, -1 for integer resolution.</param>
        ''' <returns>Dictionary of center distance vs mean value.</returns>
        ''' <remarks>We start in the middle, move down and right and always take 4 pixel symmetrical to the middle.</remarks>
        Public Shared Function Vignette(ByRef FITSSumImage(,) As UInt32, ByVal Steps As Integer) As Dictionary(Of Double, Double)
            Dim VignetPixelSum As New Dictionary(Of Double, UInt32)
            Dim VignetCount As New Dictionary(Of Double, Double)
            Dim GroupDeltaX As Integer = 1 : Dim DistX As Integer = 1
            Dim Distance As Double = Double.NaN                                                 'holds the maximum distance in the end ...
            'Move over the complete image and sum
            For DeltaX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
                Dim GroupDeltaY As Integer = 1 : Dim DistY As Integer = 1
                For DeltaY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)
                    Distance = Math.Sqrt((DistX * DistX) + (DistY * DistY))
                    Dim SampleSum As UInt32 = 0
                    SampleSum += FITSSumImage(DeltaX, DeltaY)                                  'right down
                    SampleSum += FITSSumImage(DeltaX, DeltaY - GroupDeltaY)                    'right up
                    SampleSum += FITSSumImage(DeltaX - GroupDeltaX, DeltaY)                    'left down
                    SampleSum += FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY)      'left up
                    If VignetPixelSum.ContainsKey(Distance) = False Then
                        VignetPixelSum.Add(Distance, SampleSum)
                        VignetCount.Add(Distance, 4)
                    Else
                        VignetPixelSum(Distance) += SampleSum
                        VignetCount(Distance) = CUInt(VignetCount(Distance) + 4)
                    End If
                    GroupDeltaY += 2 : DistY += 1
                Next DeltaY
                GroupDeltaX += 2 : DistX += 1
            Next DeltaX
            'Calculate the final output
            If Steps = 0 Then
                'Do not group the distance
                Dim AllKeys As New List(Of Double)(VignetPixelSum.Keys)
                For Each Entry As Double In AllKeys
                    VignetCount(Entry) = VignetPixelSum(Entry) / VignetCount(Entry)
                Next Entry
                Return VignetCount
            Else
                'Group the distance in N steps
                If Steps = -1 Then Steps = CInt(Math.Sqrt((FITSSumImage.GetUpperBound(0) * FITSSumImage.GetUpperBound(0)) + (FITSSumImage.GetUpperBound(1) * FITSSumImage.GetUpperBound(1))) / 2)
                Dim RetAccu As New Dictionary(Of Double, Ato.cSingleValueStatistics)
                Dim AllDistances As New List(Of Double)(VignetPixelSum.Keys)
                For Each SingleDistance As Double In AllDistances
                    Dim DistNorm As Double = Math.Floor(Steps * (SingleDistance / Distance))
                    Dim Value As Double = VignetPixelSum(SingleDistance) / VignetCount(SingleDistance)
                    If RetAccu.ContainsKey(DistNorm) = False Then RetAccu.Add(DistNorm, New Ato.cSingleValueStatistics(Ato.cSingleValueStatistics.eValueType.Linear))
                    RetAccu(DistNorm).AddValue(Value)
                Next SingleDistance
                Dim RetVal As New Dictionary(Of Double, Double)
                For Each Entry As Double In RetAccu.Keys
                    RetVal.Add(Entry, RetAccu(Entry).Mean)
                Next Entry
                Return RetVal
            End If
        End Function

        ''' <summary>Correct the vignette.</summary>
        Public Shared Sub CorrectVignette(ByRef FITSSumImage(,) As UInt32, ByRef VignetteCorrection As Dictionary(Of Double, Double))
            Dim GroupDeltaX As Integer = 1 : Dim DistX As Integer = 1
            Dim Distance As Integer = -1
            For DeltaX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
                Dim GroupDeltaY As Integer = 1 : Dim DistY As Integer = 1
                For DeltaY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)
                    Distance = CInt(Math.Sqrt((DistX * DistX) + (DistY * DistY)))
                    Dim Correction As Double = 1 / VignetteCorrection(Distance)
                    FITSSumImage(DeltaX, DeltaY) = CUInt(FITSSumImage(DeltaX, DeltaY) * Correction)                                                           'right down
                    FITSSumImage(DeltaX, DeltaY - GroupDeltaY) = CUInt(FITSSumImage(DeltaX, DeltaY - GroupDeltaY) * Correction)                               'right up
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY) = CUInt(FITSSumImage(DeltaX - GroupDeltaX, DeltaY) * Correction)                               'left down
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) = CUInt(FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) * Correction)   'left up
                    GroupDeltaY += 2 : DistY += 1
                Next DeltaY
                GroupDeltaX += 2 : DistX += 1
            Next DeltaX
        End Sub

    End Class

End Namespace