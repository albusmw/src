Option Explicit On
Option Strict On

'==================================================================================================================
'Atomic source file for image processing functions
'==================================================================================================================

Public Class ImageProcessing

    '''<summary>Helping structure e.g. for non-linear operations.</summary>
    Private Structure sCoordIntensity
        Public X As Integer
        Public Y As Integer
        Public Intensity As Double
        Public Shared Function Sorter(ByVal Element1 As sCoordIntensity, ByVal Element2 As sCoordIntensity) As Integer
            Return Element1.Intensity.CompareTo(Element2.Intensity)
        End Function
    End Structure

    '''<summary>Calculate a color-balanced flat image.</summary>
    '''<remarks>Color balance is done by multiplying with the median values.</remarks>
    Public Shared Sub BayerFlatBalance(ByRef Data(,) As UInt32, ByRef Stat(,) As Dictionary(Of Long, UInt32))

        Dim BayerCountX As Integer = Stat.GetUpperBound(0) + 1
        Dim BayerCountY As Integer = Stat.GetUpperBound(1) + 1
        Dim TotalChannelPixel As Long = Data.LongLength \ (BayerCountX * BayerCountY)

        'Get the median value for each bayer channel
        Dim Median(BayerCountX - 1, BayerCountY - 1) As UInt32
        Dim MedianNorm As UInt32 = UInt32.MinValue
        For Idx1 As Integer = 0 To BayerCountX - 1
            For Idx2 As Integer = 0 To BayerCountY - 1
                Dim Sum As Long = 0
                For Each Entry As UInt32 In Stat(Idx1, Idx2).Keys
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
                        Dim Pixel As UInt32 = Data(Idx1 + RGBIdx1, Idx2 + RGBIdx2)
                        Dim NewPixel As Double = Pixel * (MedianNorm / Median(RGBIdx1, RGBIdx2))
                        Data(Idx1 + RGBIdx1, Idx2 + RGBIdx2) = CUInt(NewPixel)
                    Next RGBIdx2
                Next RGBIdx1
            Next Idx2
        Next Idx1

    End Sub

    '''<summary>Make the histogram a straight line.</summary>
    '''<remarks>This is a very strange function but can auto-strech details in the image.</remarks>
    Public Shared Sub MakeHistoStraight(ByRef Data(,) As UInt16)
        Dim ToSort As New List(Of sCoordIntensity)
        For Idx1 As Integer = 0 To Data.GetUpperBound(0)
            For Idx2 As Integer = 0 To Data.GetUpperBound(1)
                Dim NewElement As sCoordIntensity
                NewElement.X = Idx1
                NewElement.Y = Idx2
                NewElement.Intensity = Data(Idx1, Idx2)
                ToSort.Add(NewElement)
            Next Idx2
        Next Idx1
        ToSort.Sort(AddressOf sCoordIntensity.Sorter)
        Dim CurrentIntense As Double = 0
        Dim IntenseStepping As Double = UInt16.MaxValue / ToSort.Count
        For Each Entry As sCoordIntensity In ToSort
            Data(Entry.X, Entry.Y) = CUShort(CurrentIntense)
            CurrentIntense += IntenseStepping
        Next Entry
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
                DiffHisto.AddTo(Diff, OneMore)
                LastEntry = Entry
            End If
        Next Entry

    End Sub

    '=========================================================================================================================
    ' VIGNETTE
    '=========================================================================================================================

    ''' <summary>Calculate the intensity over the distance from the center of the image.</summary>
    ''' <param name="ImageData">Image to run calculation on.</param>
    ''' <param name="Bins">Number of X axis steps to group - 0 for full resolution, -1 for integer resolution.</param>
    ''' <returns>Dictionary of center distance vs mean value.</returns>
    ''' <remarks>We start in the middle, move down and right and always take 4 pixel symmetrical to the middle.</remarks>
    Public Shared Function Vignette(ByRef ImageData(,) As UInt16) As Dictionary(Of Double, Double)

        Dim UInt4 As UInt32 = 4
        Dim BinSum As New Dictionary(Of Double, UInt64)
        Dim BinCount As New Dictionary(Of Double, UInt32)

        'Pre-calculate the square distances
        Dim DistSquare As New Dictionary(Of Integer, Double)
        Dim DistIdx As Integer = 1
        For CursorX As Integer = 0 To ImageData.GetUpperBound(0)
            DistSquare.Add(DistIdx, (DistIdx - 0.5) * (DistIdx - 0.5))
            DistIdx += 1
        Next CursorX

        'Move over the complete image and sum
        Dim GroupDeltaX As Integer = 1 : Dim DistXIdx As Integer = 1
        Dim DistX As Double = DistSquare(DistXIdx)
        For CursorX As Integer = (ImageData.GetUpperBound(0) \ 2) + 1 To ImageData.GetUpperBound(0)
            Dim GroupDeltaY As Integer = 1 : Dim DistYIdx As Integer = 1
            Dim DistY As Double = DistSquare(DistYIdx)
            For CursorY As Integer = (ImageData.GetUpperBound(1) \ 2) + 1 To ImageData.GetUpperBound(1)
                Dim CenterDistance As Double = (Math.Sqrt(DistX + DistY))                       'Distance from center in pixel
                Dim SampleSum As UInt32 = 0
                SampleSum += ImageData(CursorX, CursorY)                                        'right down
                SampleSum += ImageData(CursorX, CursorY - GroupDeltaY)                          'right up
                SampleSum += ImageData(CursorX - GroupDeltaX, CursorY)                          'left down
                SampleSum += ImageData(CursorX - GroupDeltaX, CursorY - GroupDeltaY)            'left up
                BinSum.AddTo(CenterDistance, SampleSum)
                BinCount.AddTo(CenterDistance, UInt4)
                GroupDeltaY += 2 : DistYIdx += 1 : DistY = DistSquare(DistYIdx)
            Next CursorY
            GroupDeltaX += 2 : DistXIdx += 1 : DistX = DistSquare(DistXIdx)
        Next CursorX

        'Calculate the final output
        Dim RetVal As New Dictionary(Of Double, Double)
        For Each Distance As Double In BinSum.Keys
            RetVal.Add(Distance, BinSum(Distance) / BinCount(Distance))
        Next Distance
        Return RetVal

    End Function

    ''' <summary>Calculate the intensity over the distance from the center of the image.</summary>
    ''' <param name="ImageData">Image to run calculation on.</param>
    ''' <param name="Steps">Number of X axis steps to group - 0 for full resolution, -1 for integer resolution.</param>
    ''' <returns>Dictionary of center distance vs mean value.</returns>
    ''' <remarks>We start in the middle, move down and right and always take 4 pixel symmetrical to the middle.</remarks>
    Public Shared Function Vignette(ByRef ImageData(,) As UInt32) As Dictionary(Of Double, Double)

        Dim UInt4 As UInt32 = 4
        Dim BinSum As New Dictionary(Of Double, UInt64)
        Dim BinCount As New Dictionary(Of Double, UInt32)

        'Pre-calculate the square distances
        Dim DistSquare As New Dictionary(Of Integer, Double)
        Dim DistIdx As Integer = 1
        For CursorX As Integer = 0 To ImageData.GetUpperBound(0)
            DistSquare.Add(DistIdx, (DistIdx - 0.5) * (DistIdx - 0.5))
            DistIdx += 1
        Next CursorX

        'Move over the complete image and sum
        Dim GroupDeltaX As Integer = 1 : Dim DistXIdx As Integer = 1
        Dim DistX As Double = DistSquare(DistXIdx)
        For CursorX As Integer = (ImageData.GetUpperBound(0) \ 2) + 1 To ImageData.GetUpperBound(0)
            Dim GroupDeltaY As Integer = 1 : Dim DistYIdx As Integer = 1
            Dim DistY As Double = DistSquare(DistYIdx)
            For CursorY As Integer = (ImageData.GetUpperBound(1) \ 2) + 1 To ImageData.GetUpperBound(1)
                Dim CenterDistance As Double = (Math.Sqrt(DistX + DistY))                       'Distance from center in pixel
                Dim SampleSum As UInt32 = 0
                SampleSum += ImageData(CursorX, CursorY)                                        'right down
                SampleSum += ImageData(CursorX, CursorY - GroupDeltaY)                          'right up
                SampleSum += ImageData(CursorX - GroupDeltaX, CursorY)                          'left down
                SampleSum += ImageData(CursorX - GroupDeltaX, CursorY - GroupDeltaY)            'left up
                BinSum.AddTo(CenterDistance, SampleSum)
                BinCount.AddTo(CenterDistance, UInt4)
                GroupDeltaY += 2 : DistYIdx += 1 : DistY = DistSquare(DistYIdx)
            Next CursorY
            GroupDeltaX += 2 : DistXIdx += 1 : DistX = DistSquare(DistXIdx)
        Next CursorX

        'Calculate the final output
        Dim RetVal As New Dictionary(Of Double, Double)
        For Each Distance As Double In BinSum.Keys
            RetVal.Add(Distance, BinSum(Distance) / BinCount(Distance))
        Next Distance
        Return RetVal

    End Function

    ''' <summary>Correct the vignette.</summary>
    ''' <returns>Number of corrected pixel.</returns>
    Public Shared Function CorrectVignette(ByRef FITSSumImage(,) As UInt16, ByRef VignetteCorrection As Dictionary(Of Double, Double)) As Integer

        Dim Corrected As Integer = 0

        Dim GroupDeltaX As Integer = 1 : Dim DistX As Integer = 1
        For DeltaX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
            Dim GroupDeltaY As Integer = 1 : Dim DistY As Integer = 1
            For DeltaY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)

                Dim CenterDistance As Double = Math.Sqrt(((DistX - 0.5) * (DistX - 0.5)) + ((DistY - 0.5) * (DistY - 0.5)))                                             'distance from center in pixel

                If VignetteCorrection.ContainsKey(CenterDistance) Then
                    Dim Correction As Double = 1 / VignetteCorrection(CenterDistance)
                    FITSSumImage(DeltaX, DeltaY) = CType(FITSSumImage(DeltaX, DeltaY) * Correction, UInt16)                                                             'right down
                    FITSSumImage(DeltaX, DeltaY - GroupDeltaY) = CType(FITSSumImage(DeltaX, DeltaY - GroupDeltaY) * Correction, UInt16)                                 'right up
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY) = CType(FITSSumImage(DeltaX - GroupDeltaX, DeltaY) * Correction, UInt16)                                 'left down
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) = CType(FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) * Correction, UInt16)     'left up
                    Corrected += 4                                                                                                                                      'number of corrected samples
                End If

                GroupDeltaY += 2 : DistY += 1
            Next DeltaY
            GroupDeltaX += 2 : DistX += 1
        Next DeltaX

        Return Corrected

    End Function

    ''' <summary>Correct the vignette.</summary>
    ''' <returns>Number of corrected pixel.</returns>
    Public Shared Function CorrectVignette(ByRef FITSSumImage(,) As UInt32, ByRef VignetteCorrection As Dictionary(Of Double, Double)) As Integer

        Dim Corrected As Integer = 0

        Dim GroupDeltaX As Integer = 1 : Dim DistX As Integer = 1
        For DeltaX As Integer = (FITSSumImage.GetUpperBound(0) \ 2) + 1 To FITSSumImage.GetUpperBound(0)
            Dim GroupDeltaY As Integer = 1 : Dim DistY As Integer = 1
            For DeltaY As Integer = (FITSSumImage.GetUpperBound(1) \ 2) + 1 To FITSSumImage.GetUpperBound(1)

                Dim CenterDistance As Double = Math.Sqrt(((DistX - 0.5) * (DistX - 0.5)) + ((DistY - 0.5) * (DistY - 0.5)))                                             'distance from center in pixel

                If VignetteCorrection.ContainsKey(CenterDistance) Then
                    Dim Correction As Double = 1 / VignetteCorrection(CenterDistance)
                    FITSSumImage(DeltaX, DeltaY) = CType(FITSSumImage(DeltaX, DeltaY) * Correction, UInt32)                                                             'right down
                    FITSSumImage(DeltaX, DeltaY - GroupDeltaY) = CType(FITSSumImage(DeltaX, DeltaY - GroupDeltaY) * Correction, UInt32)                                 'right up
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY) = CType(FITSSumImage(DeltaX - GroupDeltaX, DeltaY) * Correction, UInt32)                                 'left down
                    FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) = CType(FITSSumImage(DeltaX - GroupDeltaX, DeltaY - GroupDeltaY) * Correction, UInt32)     'left up
                    Corrected += 4                                                                                                                                      'number of corrected samples
                End If

                GroupDeltaY += 2 : DistY += 1
            Next DeltaY
            GroupDeltaX += 2 : DistX += 1
        Next DeltaX

        Return Corrected

    End Function

    '=========================================================================================================================
    ' BINNING
    '=========================================================================================================================

    '''<summary>Run a software binning.</summary>
    Public Shared Function Binning(ByRef Data(,) As UInt16, ByVal Factor As Integer) As UInt32(,)
        Dim NewWidth As Integer = CInt(Math.Floor((Data.GetUpperBound(0) + 1) / Factor))
        Dim NewHeight As Integer = CInt(Math.Floor((Data.GetUpperBound(1) + 1) / Factor))
        Dim RetVal(NewWidth - 1, NewHeight - 1) As UInt32
        Dim DataXPtr As Integer = 0
        For X As Integer = 0 To RetVal.GetUpperBound(0)
            Dim DataYPtr As Integer = 0
            For Y As Integer = 0 To RetVal.GetUpperBound(1)
                Dim NewPixel As UInt32 = 0
                For BinX As Integer = 0 To Factor - 1
                    For BinY As Integer = 0 To Factor - 1
                        NewPixel += Data(DataXPtr + BinX, DataYPtr + BinY)
                    Next BinY
                Next BinX
                RetVal(X, Y) = NewPixel
                DataYPtr += Factor
            Next Y
            DataXPtr += Factor
        Next X
        Return RetVal
    End Function

End Class