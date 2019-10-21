Option Explicit On
Option Strict On

'This class is the container for all astronomical statistics functions
'All other modules with statistics should not be used any more and put their code here
'Old modules:
' - C:\GIT\src\atomics\cStatistics.vb
'
'TODO:
' - Run multi-threaded with C:\GIT\src\atomics\cStatMultiThread.vb code
Namespace AstroNET

    Public Class Statistics

        Private Const OneUInt32 As UInt32 = CType(1, UInt32)

        Public Structure sTotalStat(Of T)
            '''<summary>Full-resolution histogram data - bayer data.</summary>
            Public BayerHistograms(,) As Dictionary(Of T, UInt32)
            '''<summary>Full-resolution histogram data - mono data.</summary>
            Public MonochromHistogram As Dictionary(Of T, UInt32)
            '''<summary>Statistics for each channel.</summary>
            Public BayerStatistics(,) As sSingleChannelStatistics(Of T)
            '''<summary>Statistics for each channel.</summary>
            Public MonochromStatsitics As sSingleChannelStatistics(Of T)
        End Structure

        '''<summary>Statistic information of one channel (RGB or total).</summary>
        Public Structure sSingleChannelStatistics(Of T)
            '''<summary>Number of total samples (pixels) in the data set.</summary>
            Public Samples As Long
            '''<summary>Maximum value occured.</summary>
            Public Max As T
            '''<summary>Minimum value occured.</summary>
            Public Min As T
            '''<summary>Value where half of the samples are below and half are above.</summary>
            Public Median As T
            '''<summary>Arithmetic mean value.</summary>
            Public Mean As Double
            '''<summary>Mean value of squared values.</summary>
            Public MeanPow2 As Double
            '''<summary>Standard deviation (calculated as in FitsWork).</summary>
            Public StdDev As Double
            '''<summary>Number of different values in the data.</summary>
            Public DifferentValueCount As Integer
            '''<summary>Init all inner variables.</summary>
            Public Function InitForShort() As sSingleChannelStatistics(Of Short)
                Dim RetVal As New sSingleChannelStatistics(Of Short)
                RetVal.Samples = 0
                RetVal.Max = Short.MinValue
                RetVal.Min = Short.MaxValue
                RetVal.Mean = 0
                RetVal.MeanPow2 = 0
                RetVal.StdDev = Double.NaN
                RetVal.DifferentValueCount = 0
                Return RetVal
            End Function
        End Structure

        '''<summary>Calculate the image statistics of the passed image data.</summary>
        Public Shared Function ImageStatistics(ByRef Data(,) As Short) As sTotalStat(Of Short)
            Dim RetVal As New sTotalStat(Of Short)
            'Calculate a bayer statistics (also for mono data - if thread-based this may even speed up ...)
            RetVal.BayerHistograms = BayerStatistics(Data, 2, 2)
            'Add all other data (mono histo and statistics)
            CalculateAllFromBayerStatistics(RetVal)
            Return RetVal
        End Function

        '''<summary>Combine 2 SingleChannelStatistics elements (e.g. to calculate the aggregated statistic for multi-frame capture).</summary>
        Public Shared Function CombineStatistics(ByVal StatA As sTotalStat(Of Short), ByVal StatB As sTotalStat(Of Short)) As sTotalStat(Of Short)
            Dim RetVal As New sTotalStat(Of Short)
            '1.) Combine to 2 histograms
            ReDim RetVal.BayerHistograms(StatA.BayerHistograms.GetUpperBound(0), StatA.BayerHistograms.GetUpperBound(1))
            For BayIdx1 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(0)
                For BayIdx2 As Integer = 0 To StatA.BayerHistograms.GetUpperBound(1)
                    'Init return bayer histogram with StatA data
                    For Each PixelValue As Short In StatA.BayerHistograms(BayIdx1, BayIdx2).Keys
                        RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, StatA.BayerHistograms(BayIdx1, BayIdx2)(PixelValue))
                    Next PixelValue
                    'Combine with StatB data
                    For Each PixelValue As Short In StatB.BayerHistograms(BayIdx1, BayIdx2).Keys
                        Dim HistoCount As UInteger = StatB.BayerHistograms(BayIdx1, BayIdx2)(PixelValue)
                        If RetVal.BayerHistograms(BayIdx1, BayIdx2).ContainsKey(PixelValue) = False Then
                            RetVal.BayerHistograms(BayIdx1, BayIdx2).Add(PixelValue, HistoCount)
                        Else
                            RetVal.BayerHistograms(BayIdx1, BayIdx2)(PixelValue) += HistoCount
                        End If
                    Next PixelValue
                Next BayIdx2
            Next BayIdx1
            CalculateAllFromBayerStatistics(RetVal)
            Return RetVal
        End Function

        '''<summary>Calculate all statistic data (mono histo and statistics) from the passed bayer statistics.</summary>
        Private Shared Sub CalculateAllFromBayerStatistics(ByRef RetVal As sTotalStat(Of Short))
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
            RetVal.MonochromStatsitics = CalcStatisticFromHistogram(RetVal.MonochromHistogram)
        End Sub

        '''<summary>Calculate the statistic data from the passed histogram data.</summary>
        Private Shared Function CalcStatisticFromHistogram(ByRef Histogram As Dictionary(Of Short, UInt32)) As sSingleChannelStatistics(Of Short)
            Dim RetVal As sSingleChannelStatistics(Of Short) : RetVal.InitForShort()
            Dim SamplesProcessed As UInt32 = 0
            'Count number of samples
            For Each KeyIdx As Short In Histogram.Keys
                RetVal.Samples += Histogram(KeyIdx)
            Next KeyIdx
            RetVal.DifferentValueCount = Histogram.Count
            For Each PixelValue As Short In Histogram.Keys
                Dim WeightCount As Double = (PixelValue * Histogram(PixelValue))
                Dim WeightPow2 As Double = (CDbl(PixelValue) * CDbl(PixelValue)) * CDbl(Histogram(PixelValue))
                SamplesProcessed += Histogram(PixelValue)
                RetVal.Mean += WeightCount
                RetVal.MeanPow2 += WeightPow2
                If PixelValue > RetVal.Max Then RetVal.Max = PixelValue
                If PixelValue < RetVal.Min Then RetVal.Min = PixelValue
                If SamplesProcessed >= RetVal.Samples \ 2 Then RetVal.Median = PixelValue
            Next PixelValue
            RetVal.StdDev = Math.Sqrt(((RetVal.MeanPow2) - ((RetVal.Mean * RetVal.Mean) / RetVal.Samples)) / (RetVal.Samples - 1))
            RetVal.Mean /= RetVal.Samples
            RetVal.MeanPow2 /= RetVal.Samples
            Return RetVal
        End Function

        '''<summary>Combine all bayer statistics to a monochromatic statistic of all pixel of the image.</summary>
        Public Shared Function CombineBayerToMonoStatistics(Of T)(ByRef BayerHistData(,) As Dictionary(Of T, UInt32)) As Dictionary(Of T, UInt32)
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
            Return RetVal
        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="XEntries">Number of different X entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<param name="YEntries">Number of different Y entries - 1 for B/W, 2 for normal RGGB, other values are exotic.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Shared Function BayerStatistics(Of T)(ByRef Data(,) As T, ByVal XEntries As Integer, ByVal YEntries As Integer) As Dictionary(Of T, UInt32)(,)

            'Count all values
            Dim RetVal(XEntries - 1, YEntries - 1) As Dictionary(Of T, UInt32)
            For Idx1 As Integer = 0 To XEntries - 1
                For Idx2 As Integer = 0 To YEntries - 1
                    RetVal(Idx1, Idx2) = BayerStatistics(Data, Idx1, XEntries, Idx2, YEntries)
                Next Idx2
            Next Idx1

            Return RetVal

        End Function

        '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
        '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
        '''<param name="OffsetX">0-based X offset where to start from.</param>
        '''<param name="OffsetY">0-based Y offset where to start from.</param>
        '''<param name="SteppingX">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
        '''<param name="SteppingY">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
        '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
        Public Shared Function BayerStatistics(Of T)(ByRef Data(,) As T, ByVal OffsetX As Integer, ByVal SteppingX As Integer, ByVal OffsetY As Integer, ByVal SteppingY As Integer) As Dictionary(Of T, UInt32)

            'Count all values
            Dim AllValues As New Dictionary(Of T, UInt32)
            For Idx1 As Integer = OffsetX To Data.GetUpperBound(0) Step SteppingX
                For Idx2 As Integer = OffsetY To Data.GetUpperBound(1) Step SteppingY
                    Dim PixelValue As T = Data(Idx1, Idx2)
                    If AllValues.ContainsKey(PixelValue) = False Then
                        AllValues.Add(PixelValue, OneUInt32)
                    Else
                        AllValues(PixelValue) += OneUInt32
                    End If
                Next Idx2
            Next Idx1

            Return SortDictionary(AllValues)

        End Function

        '================================================================================

        '''<summary>Sort the passed dictionary according to T1 (key).</summary>
        Private Shared Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2)) As Dictionary(Of T1, T2)

            'Generate a list
            Dim KeyList As New List(Of T1)
            For Each Entry As T1 In Hist.Keys
                KeyList.Add(Entry)
            Next Entry
            'Sort keys
            KeyList.Sort()
            'Re-generate dictionary
            Dim RetVal As New Dictionary(Of T1, T2)
            For Each Entry As T1 In KeyList
                RetVal.Add(Entry, Hist(Entry))
            Next Entry

            Return RetVal

        End Function

    End Class

End Namespace