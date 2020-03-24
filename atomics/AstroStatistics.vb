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

        Private Const OneUInt32 As UInt32 = CType(1, UInt32)

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
            Public BayerHistograms(,) As Collections.Generic.Dictionary(Of Int64, UInt32)
            '''<summary>Full-resolution histogram data - mono data, sorted.</summary>
            Public MonochromHistogram As Collections.Generic.Dictionary(Of Int64, UInt32)
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
            Public Shared ReadOnly Property ReportHeaderLength As Integer = 18
            '''<summary>Number of characters in the value of the report.</summary>
            Public Shared ReadOnly Property ReportValueLength As Integer = 16
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As Long
            '''<summary>Maximum value occured (value and number of pixel that have this value).</summary>
            Public Max As Collections.Generic.KeyValuePair(Of Int64, UInt32)
            '''<summary>Minimum value occured (value and number of pixel that have this value).</summary>
            Public Min As Collections.Generic.KeyValuePair(Of Int64, UInt32)
            '''<summary>Value where half of the samples are below and half are above.</summary>
            Public Median As Int64
            '''<summary>Arithmetic mean value.</summary>
            Public Mean As Double
            '''<summary>Mean value of squared values.</summary>
            Public MeanPow2 As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public StdDev As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public ReadOnly Property Variance As Double
                Get
                    Return StdDev ^ 2
                End Get
            End Property
            '''<summary>Number of different values in the data.</summary>
            Public DifferentValueCount As Integer
            '''<summary>Percentile.</summary>
            Public Percentile As Collections.Generic.Dictionary(Of Integer, Int64)
            '''<summary>Pixel value that is present the most often.</summary>
            Public Modus As Collections.Generic.KeyValuePair(Of Int64, UInt32)
            '''<summary>Init all inner variables.</summary>
            Public Shared Function InitForShort() As sSingleChannelStatistics
                Dim RetVal As New sSingleChannelStatistics
                RetVal.Samples = 0
                RetVal.Max = Nothing
                RetVal.Min = Nothing
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentValueCount = 0
                RetVal.Median = Int64.MinValue
                RetVal.Percentile = New Collections.Generic.Dictionary(Of Integer, Int64)
                RetVal.Modus = Nothing
                Return RetVal
            End Function
            '''<summary>Report of all statistics properties of the structure.</summary>
            '''<param name="DispHeader">TRUE to display the header, FALSE else.</param>
            Public Function StatisticsReport() As Collections.Generic.List(Of String)
                Dim RetVal As New Collections.Generic.List(Of String)
                RetVal.Add("Total pixel     : " & Samples.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Total pixel     : " & ((Samples / 1000000).ValRegIndep("0.0") & "M").PadLeft(ReportValueLength))
                RetVal.Add("Different values: " & DifferentValueCount.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Min value       : " & (Min.Key.ValRegIndep & " (" & Min.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Modus value     : " & (Modus.Key.ValRegIndep & " (" & Modus.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Max value       : " & (Max.Key.ValRegIndep & " (" & Max.Value.ValRegIndep & "x)").PadLeft(ReportValueLength))
                RetVal.Add("Median value    : " & Median.ValRegIndep.PadLeft(ReportValueLength))
                RetVal.Add("Mean value      : " & Format(Mean, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Standard dev.   : " & Format(StdDev, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                RetVal.Add("Variance        : " & Format(Variance, "0.000").ToString.Trim.PadLeft(ReportValueLength))
                For Each Pct As Integer In New Integer() {1, 5, 10, 25, 50, 75, 90, 95, 99}
                    If Percentile.ContainsKey(Pct) Then RetVal.Add("Percentil - " & Pct.ToString.Trim.PadLeft(2) & " %: " & Format(Percentile(Pct)).ToString.Trim.PadLeft(ReportValueLength))
                Next Pct
                Return RetVal
            End Function
        End Structure

        '''<summary>Calculate the image statistics of the passed image data.</summary>
        Public Function ImageStatistics() As sStatistics

            Dim RetVal As New sStatistics

            Dim Stopper As New System.Diagnostics.Stopwatch : Stopper.Reset() : Stopper.Start()

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
        Public Shared Function CombineStatistics(ByVal StatA As sStatistics, ByVal StatB As sStatistics) As sStatistics
            Dim RetVal As New sStatistics
            '1.) Combine to 2 histograms
            ReDim RetVal.BayerHistograms(StatA.BayerHistograms.GetUpperBound(0), StatA.BayerHistograms.GetUpperBound(1))
            For BayIdx1 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(0)
                For BayIdx2 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(1)
                    RetVal.BayerHistograms(BayIdx1, BayIdx2) = New Collections.Generic.Dictionary(Of Int64, UInteger)
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
        Private Shared Function CalcStatisticFromHistogram(ByRef Histogram As Collections.Generic.Dictionary(Of Int64, UInt32)) As sSingleChannelStatistics

            Dim RetVal As sSingleChannelStatistics = sSingleChannelStatistics.InitForShort()
            Dim AllPixelValues As Collections.Generic.List(Of Int64) = cGenerics.GetDictionaryKeys(Histogram)
            AllPixelValues.Sort()

            'Count number of samples
            For Each PixelValue As Int64 In Histogram.Keys
                RetVal.Samples += Histogram(PixelValue)
            Next PixelValue

            'Store number of different sample values
            RetVal.DifferentValueCount = Histogram.Count

            'Init statistics calculation
            Dim SumSampleCount As Long = 0
            Dim MeanSum As Double = 0
            Dim MeanPow2Sum As System.Double = 0
            RetVal.Min = New Collections.Generic.KeyValuePair(Of Int64, UInt32)(AllPixelValues(0), Histogram(AllPixelValues(0)))
            RetVal.Max = New Collections.Generic.KeyValuePair(Of Int64, UInt32)(AllPixelValues(AllPixelValues.Count - 1), Histogram(AllPixelValues(AllPixelValues.Count - 1)))
            RetVal.Modus = New Collections.Generic.KeyValuePair(Of Int64, UInt32)(AllPixelValues(0), Histogram(AllPixelValues(0)))

            'Init percentile - percentiles are writen in each bin as an incremental processing fails in fast-changing histograms
            Dim PCTInvalid As Long = Long.MinValue
            For Pct As Integer = 0 To 100
                RetVal.Percentile.Add(Pct, PCTInvalid)
            Next Pct

            'Move over the histogram
            For Each HistoX As Int64 In AllPixelValues
                Dim HistoY As UInteger = Histogram(HistoX)
                SumSampleCount += HistoY
                Dim WeightCount As Double = (CType(HistoX, Double) * CType(HistoY, Double))
                Dim WeightPow2 As Double = (CType(HistoX, Double) * CType(HistoX, Double)) * CType(HistoY, Double)
                MeanSum += WeightCount
                MeanPow2Sum += WeightPow2
                If HistoY > RetVal.Modus.Value Then RetVal.Modus = New Collections.Generic.KeyValuePair(Of Int64, UInteger)(HistoX, Histogram(HistoX))
                If SumSampleCount >= RetVal.Samples \ 2 And RetVal.Median = Int64.MinValue Then RetVal.Median = HistoX
                Dim PctIdx As Integer = CInt(100 * (SumSampleCount / RetVal.Samples))
                If RetVal.Percentile(PctIdx) = PCTInvalid Then RetVal.Percentile(PctIdx) = HistoX
            Next HistoX

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

        '''<summary>Combine all bayer statistics to a monochromatic statistic of all pixel of the image.</summary>
        Public Shared Function CombineBayerToMonoStatistics(Of T)(ByRef BayerHistData(,) As Collections.Generic.Dictionary(Of T, UInt32)) As Collections.Generic.Dictionary(Of T, UInt32)
            Dim RetVal As New Collections.Generic.Dictionary(Of T, UInt32)
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
            Return cGenerics.SortDictionary(RetVal)
        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="XEntries">Number of different X entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<param name="YEntries">Number of different Y entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Function BayerStatistics() As Collections.Generic.Dictionary(Of Int64, UInt32)(,)

            'Count all values
            Dim RetVal(1, 1) As Collections.Generic.Dictionary(Of Int64, UInt32)

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

        ''' <summary>Calculate the intensity over the distance from the center of the image.</summary>
        ''' <param name="FITSSumImage">Image to run calculation on.</param>
        ''' <param name="Steps">Number of X axis steps to group - 0 for full resolution, -1 for integer resolution.</param>
        ''' <returns>Dictionary of center distance vs mean value.</returns>
        ''' <remarks>We start in the middle, move down and right and always take 4 pixel symmetrical to the middle.</remarks>
        Public Shared Function Vignette(ByRef FITSSumImage(,) As UInt16, ByVal Steps As Integer) As Collections.Generic.Dictionary(Of Double, Double)
            Dim UInt4 As UInt32 = 4
            Dim BinSum(Steps) As UInt64
            Dim BinCount(Steps) As UInt32

            'Clear
            For Idx As Integer = 0 To Steps - 1
                BinSum(Idx) = 0 : BinCount(Idx) = 0
            Next Idx

            'Calculate the maximum distance possible from the center in X, Y and R direction
            Dim MaxDistX As Double = Double.NaN
            Dim MaxDistY As Double = Double.NaN
            Dim MaxDistance As Double = Double.NaN
            GetDistances(FITSSumImage, MaxDistX, MaxDistY, MaxDistance)

            'Move over the complete image and sum
            Dim GroupDeltaX As Integer = 1 : Dim DistXIdx As Integer = 1
            Dim DistX As Double = (DistXIdx - 0.5) * (DistXIdx - 0.5)
            For CursorX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
                Dim GroupDeltaY As Integer = 1 : Dim DistYIdx As Integer = 1
                Dim DistY As Double = (DistYIdx - 0.5) * (DistYIdx - 0.5)
                For CursorY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)
                    Dim CenterDistance As Double = (Math.Sqrt(DistX + DistY))                       'Distance from center in pixel
                    Dim DistanceBinIdx As Integer = CInt(Steps * (CenterDistance / MaxDistance))    'Index of the bin to add
                    Dim SampleSum As UInt32 = 0
                    SampleSum += FITSSumImage(CursorX, CursorY)                                     'right down
                    SampleSum += FITSSumImage(CursorX, CursorY - GroupDeltaY)                       'right up
                    SampleSum += FITSSumImage(CursorX - GroupDeltaX, CursorY)                       'left down
                    SampleSum += FITSSumImage(CursorX - GroupDeltaX, CursorY - GroupDeltaY)         'left up
                    BinSum(DistanceBinIdx) += SampleSum
                    BinCount(DistanceBinIdx) += UInt4
                    GroupDeltaY += 2 : DistYIdx += 1 : DistY = (DistYIdx - 0.5) * (DistYIdx - 0.5)
                Next CursorY
                GroupDeltaX += 2 : DistXIdx += 1 : DistX = (DistXIdx - 0.5) * (DistXIdx - 0.5)
            Next CursorX

            'Calculate the final output
            Dim RetVal As New Collections.Generic.Dictionary(Of Double, Double)
            For EntryIdx As Integer = 0 To BinSum.GetUpperBound(0)
                RetVal.Add((EntryIdx / Steps) * MaxDistance, BinSum(EntryIdx) / BinCount(EntryIdx))
            Next EntryIdx
            Return RetVal

        End Function

        ''' <summary>Calculate the intensity over the distance from the center of the image.</summary>
        ''' <param name="FITSSumImage">Image to run calculation on.</param>
        ''' <param name="Steps">Number of X axis steps to group - 0 for full resolution, -1 for integer resolution.</param>
        ''' <returns>Dictionary of center distance vs mean value.</returns>
        ''' <remarks>We start in the middle, move down and right and always take 4 pixel symmetrical to the middle.</remarks>
        Public Shared Function Vignette(ByRef FITSSumImage(,) As UInt32, ByVal Steps As Integer) As Collections.Generic.Dictionary(Of Double, Double)

            Dim UInt4 As UInt32 = 4
            Dim BinSum(Steps) As UInt64
            Dim BinCount(Steps) As UInt32

            'Clear
            For Idx As Integer = 0 To Steps - 1
                BinSum(Idx) = 0 : BinCount(Idx) = 0
            Next Idx

            'Calculate the maximum distance possible from the center in X, Y and R direction
            Dim MaxDistX As Double = Double.NaN
            Dim MaxDistY As Double = Double.NaN
            Dim MaxDistance As Double = Double.NaN
            GetDistances(FITSSumImage, MaxDistX, MaxDistY, MaxDistance)

            'Move over the complete image and sum
            Dim GroupDeltaX As Integer = 1 : Dim DistXIdx As Integer = 1
            Dim DistX As Double = (DistXIdx - 0.5) * (DistXIdx - 0.5)
            For CursorX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
                Dim GroupDeltaY As Integer = 1 : Dim DistYIdx As Integer = 1
                Dim DistY As Double = (DistYIdx - 0.5) * (DistYIdx - 0.5)
                For CursorY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)
                    Dim CenterDistance As Double = (Math.Sqrt(DistX + DistY))                       'Distance from center in pixel
                    Dim DistanceBinIdx As Integer = CInt(Steps * (CenterDistance / MaxDistance))    'Index of the bin to add
                    Dim SampleSum As UInt32 = 0
                    SampleSum += FITSSumImage(CursorX, CursorY)                                     'right down
                    SampleSum += FITSSumImage(CursorX, CursorY - GroupDeltaY)                       'right up
                    SampleSum += FITSSumImage(CursorX - GroupDeltaX, CursorY)                       'left down
                    SampleSum += FITSSumImage(CursorX - GroupDeltaX, CursorY - GroupDeltaY)         'left up
                    BinSum(DistanceBinIdx) += SampleSum
                    BinCount(DistanceBinIdx) += UInt4
                    GroupDeltaY += 2 : DistYIdx += 1 : DistY = (DistYIdx - 0.5) * (DistYIdx - 0.5)
                Next CursorY
                GroupDeltaX += 2 : DistXIdx += 1 : DistX = (DistXIdx - 0.5) * (DistXIdx - 0.5)
            Next CursorX

            'Calculate the final output
            Dim RetVal As New Collections.Generic.Dictionary(Of Double, Double)
            For EntryIdx As Integer = 0 To BinSum.GetUpperBound(0)
                RetVal.Add((EntryIdx / Steps) * MaxDistance, BinSum(EntryIdx) / BinCount(EntryIdx))
            Next EntryIdx
            Return RetVal

        End Function

        ''' <summary>Correct the vignette.</summary>
        Public Shared Sub CorrectVignette(ByRef FITSSumImage(,) As UInt32, ByRef VignetteCorrection As Collections.Generic.Dictionary(Of Double, Double))

            'Calculate the maximum distance possible from the center in X, Y and R direction
            Dim MaxDistX As Double = Double.NaN
            Dim MaxDistY As Double = Double.NaN
            Dim MaxDistance As Double = Double.NaN
            GetDistances(FITSSumImage, MaxDistX, MaxDistY, MaxDistance)
            Dim Steps As Integer = VignetteCorrection.Count - 1

            Dim GroupDeltaX As Integer = 1 : Dim DistX As Integer = 1
            For DeltaX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
                Dim GroupDeltaY As Integer = 1 : Dim DistY As Integer = 1
                For DeltaY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)

                    Dim CenterDistance As Double = Math.Sqrt(((DistX - 0.5) * (DistX - 0.5)) + ((DistY - 0.5) * (DistY - 0.5)))                                 'Distance from center in pixel
                    Dim DistanceBinIdx As Integer = CInt(Steps * (CenterDistance / MaxDistance))                                                                'Index of the bin to use for correction
                    Dim DistanceDicKey As Double = (DistanceBinIdx / Steps) * MaxDistance

                    Dim Correction As Double = 1 / VignetteCorrection(DistanceDicKey)
                    FITSSumImage(DeltaX, DeltaY) = CUInt(FITSSumImage(DeltaX, DeltaY) * Correction)                                                             'right down
                    FITSSumImage(DeltaX, DeltaY - GroupDeltaY) = CUInt(FITSSumImage(DeltaX, DeltaY - GroupDeltaY) * Correction)                                 'right up
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY) = CUInt(FITSSumImage(DeltaX - GroupDeltaX, DeltaY) * Correction)                                 'left down
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) = CUInt(FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) * Correction)     'left up
                    GroupDeltaY += 2 : DistY += 1
                Next DeltaY
                GroupDeltaX += 2 : DistX += 1
            Next DeltaX
        End Sub

        Private Shared Sub GetDistances(Of T)(ByRef FITSImage(,) As T, ByRef MaxDistX As Double, ByRef MaxDistY As Double, ByRef MaxDistance As Double)
            'Calculate the maximum distance possible from the center in X, Y and R direction
            MaxDistX = ((FITSImage.GetUpperBound(0) \ 2) + 0.5)
            MaxDistY = ((FITSImage.GetUpperBound(1) \ 2) + 0.5)
            MaxDistance = Math.Sqrt((MaxDistX * MaxDistX) + (MaxDistY * MaxDistY))
        End Sub

    End Class

End Namespace