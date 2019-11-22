Option Explicit On
Option Strict On

'''<summary>Calculate a histogram with "full X axis resolution".</summary>
'''<remarks>Should no longer be used; used AstroStatistics.vb instread ...</remarks>
Public Class cStatistics

    '================================================================================

    '''<summary>Center analyzed region.</summary>
    Public ReadOnly Property Center() As System.Drawing.PointF
        Get
            Return MyCenter
        End Get
    End Property

    '''<summary>Center analyzed region.</summary>
    Public ReadOnly Property CenterFormated() As String
        Get
            Return MyCenter.X.ToString.Trim & ":" & MyCenter.Y.ToString.Trim
        End Get
    End Property
    Private MyCenter As System.Drawing.PointF

    '''<summary>Size of the analyzed region.</summary>
    Public ReadOnly Property DetailSizeFormated() As String
        Get
            Return Width.ToString.Trim & " x " & Height.ToString.Trim
        End Get
    End Property

    Public ReadOnly Property DetailRangeFormated() As String
        Get
            Return Format(Minimum, "0.00") & "..." & Format(Maximum, "0.00")
        End Get
    End Property

    '''<summary>Width of the data calculated over.</summary>
    Public ReadOnly Property Width() As Integer
        Get
            Return MyWidth
        End Get
    End Property
    Private MyWidth As Integer = -1

    '''<summary>Height of the data calculated over.</summary>
    Public ReadOnly Property Height() As Integer
        Get
            Return MyHeight
        End Get
    End Property
    Private MyHeight As Integer = -1

    '''<summary>Total number of pixel.</summary>
    Public ReadOnly Property Pixel() As Integer
        Get
            Return MyHeight * MyWidth
        End Get
    End Property

    '''<summary>Minimum value.</summary>
    Public ReadOnly Property Minimum() As Double
        Get
            Return MyMinimum
        End Get
    End Property
    Private MyMinimum As Double = Double.MaxValue

    '''<summary>Maximum value.</summary>
    Public ReadOnly Property Maximum() As Double
        Get
            Return MyMaximum
        End Get
    End Property
    Private MyMaximum As Double = Double.MinValue

    '''<summary>Number of different values ("X axis bins").</summary>
    Public ReadOnly Property DifferentValues() As Integer
        Get
            If IsNothing(Histogram) = True Then Return 0 Else Return Histogram.Count
        End Get
    End Property

    '''<summary>25-Pertentile - 25 % of the values have a pixel amplitude below this value.</summary>
    Public ReadOnly Property Pct25() As Double
        Get
            Return MyPct25
        End Get
    End Property
    Private MyPct25 As Double = Double.NaN

    '''<summary>Median - 50 % of the values have a pixel amplitude below this value.</summary>
    Public ReadOnly Property Median() As Double
        Get
            Return MyMedian
        End Get
    End Property
    Private MyMedian As Double = Double.NaN

    '''<summary>75-Pertentile - 75 % of the values have a pixel amplitude below this value.</summary>
    Public ReadOnly Property Pct75() As Double
        Get
            Return MyPct75
        End Get
    End Property
    Private MyPct75 As Double = Double.NaN

    '''<summary>The X axis value of the histogram peak.</summary>
    Public ReadOnly Property HistoPeakPos() As Double
        Get
            Return MyHistoPeakPos
        End Get
    End Property
    Private MyHistoPeakPos As Double = Double.NaN

    '''<summary>The Y axis value of the histogram peak.</summary>
    Public ReadOnly Property HistoPeakCount() As Integer
        Get
            Return MyHistoPeakCount
        End Get
    End Property
    Private MyHistoPeakCount As Integer = 0

    '''<summary>Mean pixel value.</summary>
    Public ReadOnly Property Mean() As Double
        Get
            Return MyMean
        End Get
    End Property
    Private MyMean As Double = Double.NaN

    '''<summary>Standard deviation value.</summary>
    Public ReadOnly Property StdDev() As Double
        Get
            Return MyStdDev
        End Get
    End Property
    Private MyStdDev As Double = Double.NaN

    '================================================================================

    Public Histogram As Dictionary(Of Double, Integer)

    '================================================================================

    '''<summary>Calculate statistics on the given area.</summary>
    Public Sub Calculate(ByRef Data(,) As Double)
        Calculate(Data, 0, Data.GetUpperBound(0), 0, Data.GetUpperBound(1))
    End Sub

    '''<summary>Calculate statistics on the given area.</summary>
    Public Sub Calculate(ByRef Data(,) As Double, ByVal X_left As Integer, ByVal X_right As Integer, ByVal Y_top As Integer, ByVal Y_bottom As Integer)

        Dim Pixels As Integer = (X_right - X_left + 1) * (Y_bottom - Y_top + 1)

        'Set area that was calculated
        MyWidth = X_right - X_left + 1
        MyHeight = Y_bottom - Y_top + 1
        Pixels = MyWidth * MyHeight

        'Set center
        MyCenter = New System.Drawing.PointF(CSng(X_left + (MyWidth / 2)), CSng(Y_top + (MyHeight / 2)))

        'Reset
        MyMaximum = Double.MinValue
        MyMinimum = Double.MaxValue
        MyMean = 0
        Dim MeanPow2 As Double = 0

        'Build the histogramm
        Histogram = New Dictionary(Of Double, Integer)
        For Idx1 As Integer = X_left To X_right
            For Idx2 As Integer = Y_top To Y_bottom
                Dim Value As Double = Data(Idx1, Idx2)
                If Histogram.ContainsKey(Value) Then Histogram(Value) += 1 Else Histogram.Add(Value, 1)
                MyMean += Value
                MeanPow2 += Value * Value
                If Value > MyMaximum Then MyMaximum = Value
                If Value < MyMinimum Then MyMinimum = Value
            Next Idx2
        Next Idx1
        Histogram = SortDictionary(Histogram)

        MyMean /= Pixels
        MeanPow2 /= Pixels
        MyStdDev = Math.Sqrt(MeanPow2 - (MyMean * MyMean))

        'Get percentiles
        Dim CountedPixel As Integer = 0
        MyPct25 = Double.NaN
        MyMedian = Double.NaN
        MyPct75 = Double.NaN
        MyHistoPeakCount = 0
        MyHistoPeakPos = Double.NaN
        For Each Entry As Double In Histogram.Keys
            CountedPixel += Histogram(Entry)
            If Histogram(Entry) > MyHistoPeakCount Then
                MyHistoPeakCount = Histogram(Entry)
                MyHistoPeakPos = Entry
            End If
            If CountedPixel >= Pixels * 0.25 And Double.IsNaN(MyPct25) = True Then MyPct25 = Entry
            If CountedPixel >= Pixels * 0.5 And Double.IsNaN(MyMedian) = True Then MyMedian = Entry
            If CountedPixel >= Pixels * 0.75 And Double.IsNaN(MyPct75) = True Then MyPct75 = Entry
        Next Entry

    End Sub


    '================================================================================

    '''<summary>Calculate statistics on the given area.</summary>
    Public Sub Calculate(ByRef Data(,) As Int32)
        Calculate(Data, 0, Data.GetUpperBound(0), 0, Data.GetUpperBound(1))
    End Sub

    '''<summary>Calculate statistics on the given area.</summary>
    Public Sub Calculate(ByRef Data(,) As Int32, ByVal X_left As Integer, ByVal X_right As Integer, ByVal Y_top As Integer, ByVal Y_bottom As Integer)

        Dim Pixels As Integer = (X_right - X_left + 1) * (Y_bottom - Y_top + 1)

        'Set area that was calculated
        MyWidth = X_right - X_left + 1
        MyHeight = Y_bottom - Y_top + 1
        Pixels = MyWidth * MyHeight

        'Set center
        MyCenter = New System.Drawing.PointF(CSng(X_left + (MyWidth / 2)), CSng(Y_top + (MyHeight / 2)))

        'Reset
        MyMaximum = Double.MinValue
        MyMinimum = Double.MaxValue
        MyMean = 0
        Dim MeanPow2 As Double = 0

        'Build the histogramm
        Histogram = New Dictionary(Of Double, Integer)
        For Idx1 As Integer = X_left To X_right
            For Idx2 As Integer = Y_top To Y_bottom
                Dim Value As Double = Data(Idx1, Idx2)
                If Histogram.ContainsKey(Value) Then Histogram(Value) += 1 Else Histogram.Add(Value, 1)
                MyMean += Value
                MeanPow2 += Value * Value
                If Value > MyMaximum Then MyMaximum = Value
                If Value < MyMinimum Then MyMinimum = Value
            Next Idx2
        Next Idx1
        Histogram = SortDictionary(Histogram)

        MyMean /= Pixels
        MeanPow2 /= Pixels
        MyStdDev = Math.Sqrt(MeanPow2 - (MyMean * MyMean))

        'Get percentiles
        Dim CountedPixel As Integer = 0
        MyPct25 = Double.NaN
        MyMedian = Double.NaN
        MyPct75 = Double.NaN
        MyHistoPeakCount = 0
        MyHistoPeakPos = Double.NaN
        For Each Entry As Double In Histogram.Keys
            CountedPixel += Histogram(Entry)
            If Histogram(Entry) > MyHistoPeakCount Then
                MyHistoPeakCount = Histogram(Entry)
                MyHistoPeakPos = Entry
            End If
            If CountedPixel >= Pixels * 0.25 And Double.IsNaN(MyPct25) = True Then MyPct25 = Entry
            If CountedPixel >= Pixels * 0.5 And Double.IsNaN(MyMedian) = True Then MyMedian = Entry
            If CountedPixel >= Pixels * 0.75 And Double.IsNaN(MyPct75) = True Then MyPct75 = Entry
        Next Entry

    End Sub

    '================================================================================

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
