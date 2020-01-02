Option Explicit On
Option Strict On

'''<summary>Functions to calculate or adjust bayer matrix data.</summary>
'''<remarks>This functions are old ones copied and need a review ...</remarks>
Public Class BayerFunctions

    Private Const OneUInt32 As UInt32 = CType(1, UInt32)

    '''<summary>Calculate basic bayer statistics on the passed data matrix.</summary>
    '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
    '''<param name="OffsetX">0-based X offset where to start from.</param>
    '''<param name="OffsetY">0-based Y offset where to start from.</param>
    '''<param name="SteppingX">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<param name="SteppingY">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<returns>A sorted dictionary which contains all found values of type T in the Data matrix and its count.</returns>
    <Obsolete("This class should be inspected ...", True)>
    Public Shared Function BayerStatistics(Of T)(ByRef Data(,) As T) As Dictionary(Of T, UInt32)(,)

        'Count all values
        Dim RetVal(1, 1) As Dictionary(Of T, UInt32)
        For Idx1 As Integer = 0 To 1
            For Idx2 As Integer = 0 To 1
                RetVal(Idx1, Idx2) = BayerStatistics(Data, Idx1, 2, Idx2, 2)
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
    <Obsolete("This class should be inspected ...", True)>
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

        Return cGenerics.SortDictionary(AllValues)

    End Function

    '''<summary>Get only one bayer channel by removing empty rows and columns.</summary>
    '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
    '''<param name="OffsetX">0-based X offset where to start from.</param>
    '''<param name="OffsetY">0-based Y offset where to start from.</param>
    '''<param name="SteppingX">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<param name="SteppingY">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<returns>The linear interpolated data.</returns>
    <Obsolete("This class should be inspected ...", True)>
    Public Shared Function CompressBayerChannel(ByRef Data(,) As Int32, ByVal OffsetX As Integer, ByVal SteppingX As Integer, ByVal OffsetY As Integer, ByVal SteppingY As Integer) As Int32(,)

        Dim OutData(Data.GetUpperBound(0) \ 2, Data.GetUpperBound(1) \ 2) As Int32

        'Copy original data
        Dim Ptr1 As Integer = 0
        For Idx1 As Integer = OffsetX To Data.GetUpperBound(0) Step SteppingX
            Dim Ptr2 As Integer = 0
            For Idx2 As Integer = OffsetY To Data.GetUpperBound(1) Step SteppingY
                OutData(Ptr1, Ptr2) = Data(Idx1, Idx2)
                Ptr2 += 1
            Next Idx2
            Ptr1 += 1
        Next Idx1

        Return OutData

    End Function

    '''<summary>Interpolate the bayer channel, means that we e.g. only get the R channel but in full-resolution.</summary>
    '''<param name="Data">Matrix of data - 2D matrix what contains the raw sensor data.</param>
    '''<param name="OffsetX">0-based X offset where to start from.</param>
    '''<param name="OffsetY">0-based Y offset where to start from.</param>
    '''<param name="SteppingX">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<param name="SteppingY">Step size in X direction - typically 2 for a normal RGGB bayer matrix.</param>
    '''<returns>The linear interpolated data.</returns>
    <Obsolete("This class should be inspected ...", True)>
    Public Shared Function InterpolateBayerChannel(ByRef Data(,) As Int32, ByVal OffsetX As Integer, ByVal SteppingX As Integer, ByVal OffsetY As Integer, ByVal SteppingY As Integer) As Int32(,)

        Dim OutData(Data.GetUpperBound(0), Data.GetUpperBound(1)) As Int32

        'Copy original data
        For Idx1 As Integer = OffsetX To Data.GetUpperBound(0) Step SteppingX
            For Idx2 As Integer = OffsetY To Data.GetUpperBound(1) Step SteppingY
                OutData(Idx1, Idx2) = Data(Idx1, Idx2)
            Next Idx2
        Next Idx1

        'R channel
        If OffsetX = 0 And OffsetY = 0 Then
            '1.) Fill the diagonal values in the odd rows (from top left and right / bottom left and right)
            For Idx1 As Integer = 1 To Data.GetUpperBound(0) - 1 Step SteppingX
                For Idx2 As Integer = 1 To Data.GetUpperBound(1) - 1 Step SteppingY
                    Dim NewVal As Int32 = 0
                    NewVal += OutData(Idx1 - 1, Idx2 - 1)
                    NewVal += OutData(Idx1 - 1, Idx2 + 1)
                    NewVal += OutData(Idx1 + 1, Idx2 - 1)
                    NewVal += OutData(Idx1 + 1, Idx2 + 1)
                    OutData(Idx1, Idx2) = NewVal \ 4
                Next Idx2
            Next Idx1
            '2.) Fill the missing values in the odd rows (from top / bottom / left / right)
            For Idx1 As Integer = 2 To Data.GetUpperBound(0) - 1 Step SteppingX
                For Idx2 As Integer = 1 To Data.GetUpperBound(1) - 1 Step SteppingY
                    Dim NewVal As Int32 = 0
                    NewVal += OutData(Idx1, Idx2 + 1)           'top
                    NewVal += OutData(Idx1, Idx2 - 1)           'bottom
                    NewVal += OutData(Idx1 - 1, Idx2)           'left
                    NewVal += OutData(Idx1 + 1, Idx2)           'right
                    OutData(Idx1, Idx2) = NewVal \ 4
                Next Idx2
            Next Idx1
            '3.) Fill all even values (from top / bottom / left / right)
            For Idx1 As Integer = 1 To Data.GetUpperBound(0) - 1 Step SteppingX
                For Idx2 As Integer = 2 To Data.GetUpperBound(1) - 1 Step SteppingY
                    Dim NewVal As Int32 = 0
                    NewVal += OutData(Idx1, Idx2 + 1)           'top
                    NewVal += OutData(Idx1, Idx2 - 1)           'bottom
                    NewVal += OutData(Idx1 - 1, Idx2)           'left
                    NewVal += OutData(Idx1 + 1, Idx2)           'right
                    OutData(Idx1, Idx2) = NewVal \ 4
                Next Idx2
            Next Idx1
        End If

        Return OutData

    End Function

    '''<summary>Equalize the bayer data in the passed image.</summary>
    '''<param name="ImageData">Data to equalize.</param>
    '''<param name="ImageStatistics">Bayer statistics data.</param>
    '''<returns>List of calculation details as string log.</returns>
    Public Shared Function EqualizeBayerChannels(ByRef ImageData(,) As UInt32, ByVal ImageStatistics As AstroNET.Statistics.sStatistics) As List(Of String)

        Dim RetVal As New List(Of String)

        'Calculate the maximum median value
        Dim ReferenceValue As Double = Double.MinValue
        For Idx1 As Integer = 0 To 1
            For Idx2 As Integer = 0 To 1
                If ImageStatistics.BayerStatistics(Idx1, Idx2).Mean > ReferenceValue Then ReferenceValue = ImageStatistics.BayerStatistics(Idx1, Idx2).Mean
            Next Idx2
        Next Idx1
        RetVal.Add("Max mean: " & ReferenceValue.ToString.Trim)

        'Calculate the normalization data (multiplier to get the same maximum median in all channels)
        Dim Norm(1, 1) As Double
        For Idx1 As Integer = 0 To 1
            For Idx2 As Integer = 0 To 1
                Norm(Idx1, Idx2) = ReferenceValue / ImageStatistics.BayerStatistics(Idx1, Idx2).Mean
                RetVal.Add("Norm[" & Idx1.ToString.Trim & ":" & Idx2.ToString.Trim & "]           : " & Norm(Idx1, Idx2).ToString.Trim)
                RetVal.Add("Max [" & Idx1.ToString.Trim & ":" & Idx2.ToString.Trim & "] after norm: " & CType(Math.Round(ImageStatistics.BayerStatistics(Idx1, Idx2).Max * Norm(Idx1, Idx2)), UInt32).ToString.Trim)
            Next Idx2
        Next Idx1

        'Apply the normalization
        For IdxX As Integer = 0 To ImageData.GetUpperBound(0) Step 2
            For IdxY As Integer = 0 To ImageData.GetUpperBound(1) Step 2
                For Idx1 As Integer = 0 To 1
                    For Idx2 As Integer = 0 To 1
                        ImageData(IdxX + Idx1, IdxY + Idx2) = CType(Math.Round(ImageData(IdxX + Idx1, IdxY + Idx2) * Norm(Idx1, Idx2)), UInt32)
                    Next Idx2
                Next Idx1
            Next IdxY
        Next IdxX

        Return RetVal

    End Function

    '''<summary>Equalize the bayer data in the passed image.</summary>
    <Obsolete("This class should be inspected ...", True)>
    Public Shared Sub EqualizeChannels(ByRef ImageData(,) As Short)

        Dim Mean_00 As Double = 0
        Dim Mean_01 As Double = 0
        Dim Mean_10 As Double = 0
        Dim Mean_11 As Double = 0

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                Mean_00 += ImageData(Idx1, Idx2)
                Mean_01 += ImageData(Idx1, Idx2 + 1)
                Mean_10 += ImageData(Idx1 + 1, Idx2)
                Mean_11 += ImageData(Idx1 + 1, Idx2 + 1)
            Next Idx2
        Next Idx1

        Mean_00 /= (ImageData.LongLength / 4)
        Mean_01 /= (ImageData.LongLength / 4)
        Mean_10 /= (ImageData.LongLength / 4)
        Mean_11 /= (ImageData.LongLength / 4)

        Dim NewMean As Double = (Mean_00 + Mean_01 + Mean_10 + Mean_11) / 4

        Mean_00 = NewMean / Mean_00
        Mean_01 = NewMean / Mean_01
        Mean_10 = NewMean / Mean_10
        Mean_11 = NewMean / Mean_11

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                ImageData(Idx1, Idx2) = CShort(ImageData(Idx1, Idx2) * Mean_00)
                ImageData(Idx1, Idx2 + 1) = CShort(ImageData(Idx1, Idx2 + 1) * Mean_01)
                ImageData(Idx1 + 1, Idx2) = CShort(ImageData(Idx1 + 1, Idx2) * Mean_10)
                ImageData(Idx1 + 1, Idx2 + 1) = CShort(ImageData(Idx1 + 1, Idx2 + 1) * Mean_11)
            Next Idx2
        Next Idx1

    End Sub

    '''<summary>Equalize the bayer data in the passed image.</summary>
    <Obsolete("This class should be inspected ...", False)>
    Public Shared Sub EqualizeChannels(ByRef ImageData(,) As Int32)

        Dim Mean_00 As Double = 0
        Dim Mean_01 As Double = 0
        Dim Mean_10 As Double = 0
        Dim Mean_11 As Double = 0

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                Mean_00 += ImageData(Idx1, Idx2)
                Mean_01 += ImageData(Idx1, Idx2 + 1)
                Mean_10 += ImageData(Idx1 + 1, Idx2)
                Mean_11 += ImageData(Idx1 + 1, Idx2 + 1)
            Next Idx2
        Next Idx1

        Mean_00 /= (ImageData.LongLength / 4)
        Mean_01 /= (ImageData.LongLength / 4)
        Mean_10 /= (ImageData.LongLength / 4)
        Mean_11 /= (ImageData.LongLength / 4)

        Dim NewMean As Double = (Mean_00 + Mean_01 + Mean_10 + Mean_11) / 4

        Mean_00 = NewMean / Mean_00
        Mean_01 = NewMean / Mean_01
        Mean_10 = NewMean / Mean_10
        Mean_11 = NewMean / Mean_11

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                ImageData(Idx1, Idx2) = CType(ImageData(Idx1, Idx2) * Mean_00, Int32)
                ImageData(Idx1, Idx2 + 1) = CType(ImageData(Idx1, Idx2 + 1) * Mean_01, Int32)
                ImageData(Idx1 + 1, Idx2) = CType(ImageData(Idx1 + 1, Idx2) * Mean_10, Int32)
                ImageData(Idx1 + 1, Idx2 + 1) = CType(ImageData(Idx1 + 1, Idx2 + 1) * Mean_11, Int32)
            Next Idx2
        Next Idx1

    End Sub

    '''<summary>Equalize the bayer data in the passed image.</summary>
    <Obsolete("This class should be inspected ...", False)>
    Public Shared Sub EqualizeChannels(ByRef ImageData(,) As Double)

        Dim Mean_00 As Double = 0
        Dim Mean_01 As Double = 0
        Dim Mean_10 As Double = 0
        Dim Mean_11 As Double = 0

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                Mean_00 += ImageData(Idx1, Idx2)
                Mean_01 += ImageData(Idx1, Idx2 + 1)
                Mean_10 += ImageData(Idx1 + 1, Idx2)
                Mean_11 += ImageData(Idx1 + 1, Idx2 + 1)
            Next Idx2
        Next Idx1

        Mean_00 /= (ImageData.LongLength / 4)
        Mean_01 /= (ImageData.LongLength / 4)
        Mean_10 /= (ImageData.LongLength / 4)
        Mean_11 /= (ImageData.LongLength / 4)

        Dim NewMean As Double = (Mean_00 + Mean_01 + Mean_10 + Mean_11) / 4

        Mean_00 = NewMean / Mean_00
        Mean_01 = NewMean / Mean_01
        Mean_10 = NewMean / Mean_10
        Mean_11 = NewMean / Mean_11

        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step 2
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step 2
                ImageData(Idx1, Idx2) *= Mean_00
                ImageData(Idx1, Idx2 + 1) *= Mean_01
                ImageData(Idx1 + 1, Idx2) *= Mean_10
                ImageData(Idx1 + 1, Idx2 + 1) *= Mean_11
            Next Idx2
        Next Idx1

    End Sub

    <Obsolete("This class should be inspected ...", False)>
    Public Shared Function ClusterBrightness(ByRef ImageData(,) As Double, ByRef BaseLine() As Double, ByRef Traces()() As Double) As Double

        Dim BayerX As Integer = 2
        Dim BayerY As Integer = 2
        Dim TraceLength As Integer = CInt(ImageData.LongLength \ (BayerX * BayerY))

        'Init for a 2x2 bayer configuration
        Dim BaseLineList As New List(Of Double)
        ReDim Traces((BayerX * BayerY) - 1)

        For Idx As Integer = 0 To Traces.GetUpperBound(0)
            ReDim Traces(Idx)(TraceLength - 1)
        Next Idx

        'Generate data
        Dim Ptr As Integer = 0
        Dim Max As Double = Double.MinValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0) - 1 Step BayerX
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1) - 1 Step BayerY
                Dim Brightness As Double = 0
                Brightness += ImageData(Idx1, Idx2) : Traces(0)(Ptr) = (ImageData(Idx1, Idx2))
                Brightness += ImageData(Idx1, Idx2 + 1) : Traces(1)(Ptr) = (ImageData(Idx1, Idx2 + 1))
                Brightness += ImageData(Idx1 + 1, Idx2) : Traces(2)(Ptr) = (ImageData(Idx1 + 1, Idx2))
                Brightness += ImageData(Idx1 + 1, Idx2 + 1) : Traces(3)(Ptr) = (ImageData(Idx1 + 1, Idx2 + 1))
                Brightness /= 4
                BaseLineList.Add(Brightness)
                If Brightness > Max Then Max = Brightness
                Ptr += 1
            Next Idx2
        Next Idx1

        BaseLine = BaseLineList.ToArray

        Return Max

    End Function

End Class
