Option Explicit On
Option Strict On

'''<summary>Functions to calculate or adjust bayer matrix data.</summary>
Public Class BayerFunctions

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
