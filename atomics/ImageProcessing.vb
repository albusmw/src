Option Explicit On
Option Strict On

'==================================================================================================================
'Atomic source file for image processing functions
'==================================================================================================================

Public Class ImageProcessing   

    '''<summary>Calculate a color-balanced flat image.</summary>
    '''<remarks>Color balance is done by multiplying with the median values.</remarks>
    Public Shared Sub BayerFlatBalance(ByRef Data(,) As Int32, ByRef Stat(,) As Dictionary(Of Int32, UInt32))

        Dim BayerCountX As Integer = Stat.GetUpperBound(0) + 1
        Dim BayerCountY As Integer = Stat.GetUpperBound(1) + 1
        Dim TotalChannelPixel As Long = Data.LongLength \ (BayerCountX * BayerCountY)

        'Get the median value for each bayer channel
        Dim Median(BayerCountX - 1, BayerCountY - 1) As Int32
        Dim MedianNorm As Int32 = Int32.MinValue
        For Idx1 As Integer = 0 To BayerCountX - 1
            For Idx2 As Integer = 0 To BayerCountY - 1
                Dim Sum As Long = 0
                For Each Entry As Int32 In Stat(Idx1, Idx2).Keys
                    Sum += Stat(Idx1, Idx2)(Entry)
                    If Sum >= TotalChannelPixel \ 2 Then
                        Median(Idx1, Idx2) = Entry
                        If Median(Idx1, Idx2) > MedianNorm Then MedianNorm = Median(Idx1, Idx2)
                        Exit For
                    End If
                Next Entry
            Next Idx2
        Next Idx1

        'Correct all bayer channels to match the histogram with the maximum value and also correct the histogram data
        Dim NewStat(Stat.GetUpperBound(0), Stat.GetUpperBound(1)) As Dictionary(Of Int32, UInt32)
        For Idx1 As Integer = 0 To Data.GetUpperBound(0) - 1 Step BayerCountX
            For Idx2 As Integer = 0 To Data.GetUpperBound(1) - 1 Step BayerCountY
                For RGBIdx1 As Integer = 0 To BayerCountX - 1
                    For RGBIdx2 As Integer = 0 To BayerCountY - 1
                        Dim Pixel As Int32 = Data(Idx1 + RGBIdx1, Idx2 + RGBIdx2)
                        Dim NewPixel As Double = Pixel * (MedianNorm / Median(RGBIdx1, RGBIdx2))
                        Data(Idx1 + RGBIdx1, Idx2 + RGBIdx2) = CInt(NewPixel)
                    Next RGBIdx2
                Next RGBIdx1
            Next Idx2
        Next Idx1

    End Sub

    '''<summary>Calculate the number of all samples taken.</summary>
    Public Shared Function HistoCount(ByRef Histo As Dictionary(Of Int32, UInt32)) As Long
        Dim Count As Long = 0
        For Each Entry As Int32 In Histo.Keys
            Count += Histo(Entry)
        Next Entry
        Return Count
    End Function

    '''<summary>Calculate the mean value of the given histogramm.</summary>
    Public Shared Function HistoMean(ByRef Histo As Dictionary(Of Int32, UInt32)) As Double
        Dim Sum As Double = 0
        Dim Count As Long = 0
        For Each Entry As Int32 In Histo.Keys
            Count += Histo(Entry)
            Sum += Entry * Histo(Entry)
        Next Entry
        Return Sum / Count
    End Function

    '''<summary>Calculate basic histogramm parameters.</summary>
    Public Shared Sub HistogramParameters(ByRef Histo As Dictionary(Of Int32, UInt32), ByRef DiffHisto As Dictionary(Of Int32, UInt32))

        Dim OneMore As UInt32 = CType(1, UInt32)

        'Differential statistics
        DiffHisto = New Dictionary(Of Int32, UInt32)
        Dim FirstOne As Boolean = True
        Dim LastEntry As Int32 = Int32.MinValue
        For Each Entry As Int32 In Histo.Keys
            If FirstOne = True Then
                LastEntry = Entry : FirstOne = False
            Else
                Dim Diff As Int32 = Entry - LastEntry
                If DiffHisto.ContainsKey(Diff) = False Then
                    DiffHisto.Add(Diff, 1)
                Else
                    DiffHisto(Diff) += OneUInt32
                End If
                LastEntry = Entry
            End If
        Next Entry

    End Sub

    '''<summary>Sort the passed dictionary according to T1 (key).</summary>
    Public Shared Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2)) As Dictionary(Of T1, T2)

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