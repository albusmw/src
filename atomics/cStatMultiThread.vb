Option Explicit On
Option Strict On

'''<summary>Type for results of ADU value counts.</summary>
Imports ADUCount = System.UInt64
'''<summary>Type for interger ADU statistics.</summary>
Imports ADUFixed = System.Int64

Public Class cStatMultiThread
    Public Const OneUInt64 As ADUCount = CType(1, ADUCount)
    '''<summary>Object for each thread.</summary>
    Public Class cStatObjFixed
        Friend NAXIS3 As Integer = -1
        Friend XOffset As Integer = -1
        Friend YOffset As Integer = -1
        Friend HistDataBayer As New Dictionary(Of ADUFixed, ADUCount)
        Friend Done As Boolean = False
    End Class
End Class

'=================================================================================================================================
' UInt16
'=================================================================================================================================

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
'''<remarks>Calculation is done by buidling a vector with all possible entries (only 2^16 length).</remarks>
Public Class cStatMultiThread_UInt16

    Public Structure sImgData_UInt16
        Public Data(,) As UInt16
        Public ReadOnly Property Length() As Long
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.LongLength
            End Get
        End Property
        Public ReadOnly Property NAXIS1() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(0) + 1
            End Get
        End Property
        Public ReadOnly Property NAXIS2() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(1) + 1
            End Get
        End Property
    End Structure

    '''<summary>The real image data - 3 is the maximum NAXIS3 value for e.g. LRGB images.</summary>
    Public ImageData(3) As sImgData_UInt16

    '''<summary>Load UInt32 data to the internal UInt16 structure.</summary>
    Public Sub LoadImageData(ByVal Data(,) As UInt32)
        ReDim ImageData(0).Data(Data.GetUpperBound(0), Data.GetUpperBound(1))
        Threading.Tasks.Parallel.For(0, Data.GetUpperBound(0), Sub(Idx1 As Integer)
                                                                   For Idx2 As Integer = 0 To Data.GetUpperBound(1)
                                                                       ImageData(0).Data(Idx1, Idx2) = CUShort(Data(Idx1, Idx2))
                                                                   Next Idx2
                                                               End Sub)
    End Sub

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub RunHistoCalc(ByVal NAXIS3 As Integer, ByRef Results(,) As cStatMultiThread.cStatObjFixed)

        'Data are processed
        Dim StObj(3) As cStatMultiThread.cStatObjFixed
        For Idx As Integer = 0 To StObj.GetUpperBound(0)
            StObj(Idx) = New cStatMultiThread.cStatObjFixed
        Next Idx
        StObj(0).NAXIS3 = NAXIS3 : StObj(0).XOffset = 0 : StObj(0).YOffset = 0
        StObj(1).NAXIS3 = NAXIS3 : StObj(1).XOffset = 0 : StObj(1).YOffset = 1
        StObj(2).NAXIS3 = NAXIS3 : StObj(2).XOffset = 1 : StObj(2).YOffset = 0
        StObj(3).NAXIS3 = NAXIS3 : StObj(3).XOffset = 1 : StObj(3).YOffset = 1

        'Start all threads
        For Each Slice As cStatMultiThread.cStatObjFixed In StObj
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
        Next Slice

        'Join all threads
        Do
            'System.Threading.Thread.Sleep(1)
            Dim AllDone As Boolean = True
            For Each Slice As cStatMultiThread.cStatObjFixed In StObj
                If Slice.Done = False Then
                    AllDone = False : Exit For
                End If
            Next Slice
            If AllDone Then Exit Do
        Loop Until 1 = 0

        'Collect all results
        ReDim Results(1, 1)
        Results(0, 0) = StObj(0)
        Results(0, 1) = StObj(1)
        Results(1, 0) = StObj(2)
        Results(1, 1) = StObj(3)

    End Sub

    '''<summary>Histogramm calculation itself - the histogram of one bayer channel is calculated.</summary>
    '''<remarks>The histogramm is calculated by forming a vector with all possible entries and using the pixel value direct as index.</remarks>
    Private Sub HistoCalc(ByVal Arguments As Object)

        Dim StateObj As cStatMultiThread.cStatObjFixed = CType(Arguments, cStatMultiThread.cStatObjFixed)
        StateObj.Done = False

        'Init count object with 0
        Dim HistCount(UInt16.MaxValue) As ADUCount
        For Idx As Integer = 0 To HistCount.GetUpperBound(0)
            HistCount(Idx) = 0
        Next Idx

        'Count one bayer part
        For IdxX As Integer = StateObj.XOffset To ImageData(StateObj.NAXIS3).Data.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData(StateObj.NAXIS3).Data.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                HistCount(ImageData(StateObj.NAXIS3).Data(IdxX, IdxY)) += cStatMultiThread.OneUInt64
            Next IdxY
        Next IdxX

        'Form return value with maximum value range ADHFixed
        StateObj.HistDataBayer = New Dictionary(Of ADUFixed, ADUCount)
        For Idx As Integer = 0 To HistCount.GetUpperBound(0)
            If HistCount(Idx) > 0 Then StateObj.HistDataBayer.Add(Idx, HistCount(Idx))
            If Idx = HistCount.GetUpperBound(0) Then Exit For
        Next Idx

        StateObj.Done = True

    End Sub

    '''<summary>Get a list of pixel above a certain value.</summary>
    '''<param name="ValueAbove">Value (included) to search for.</param>
    Public Function GetAbove(ByVal ValueAbove As UInt16) As Dictionary(Of UInt16, List(Of Drawing.Point))

        'Find top 1% of values and create a dictionary for the values and all pixel with this value
        Dim RetVal As New Dictionary(Of UInt16, List(Of Drawing.Point))
        With ImageData(0)
            For Idx1 As Integer = 0 To .NAXIS1 - 1
                For Idx2 As Integer = 0 To .NAXIS2 - 1
                    If .Data(Idx1, Idx2) >= ValueAbove Then
                        If RetVal.ContainsKey(.Data(Idx1, Idx2)) = False Then
                            RetVal.Add(.Data(Idx1, Idx2), New List(Of Drawing.Point)({New Drawing.Point(Idx1, Idx2)}))
                        Else
                            RetVal(.Data(Idx1, Idx2)).Add(New Drawing.Point(Idx1, Idx2))
                        End If
                    End If
                Next Idx2
            Next Idx1
        End With

        RetVal = RetVal.SortDictionaryInverse
        Return RetVal

    End Function

End Class

'=================================================================================================================================
' UInt32
'=================================================================================================================================

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
'''<remarks>Calculation as for UInt16 is not possible as a vector with all entries would be 2^32 entries long.</remarks>
Public Class cStatMultiThread_UInt32

    Public Structure sImgData_UInt32
        Public Data(,) As UInt32
        Public ReadOnly Property Length() As Long
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.LongLength
            End Get
        End Property
        Public ReadOnly Property NAXIS1() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(0) + 1
            End Get
        End Property
        Public ReadOnly Property NAXIS2() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(1) + 1
            End Get
        End Property
    End Structure

    '''<summary>The real image data.</summary>
    Public ImageData(3) As sImgData_UInt32

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub RunHistoCalc(ByVal NAXIS3 As Integer, ByRef Results(,) As cStatMultiThread.cStatObjFixed)

        'Data are processed
        Dim StObj(3) As cStatMultiThread.cStatObjFixed
        For Idx As Integer = 0 To StObj.GetUpperBound(0)
            StObj(Idx) = New cStatMultiThread.cStatObjFixed
        Next Idx
        StObj(0).NAXIS3 = NAXIS3 : StObj(0).XOffset = 0 : StObj(0).YOffset = 0
        StObj(1).NAXIS3 = NAXIS3 : StObj(1).XOffset = 0 : StObj(1).YOffset = 1
        StObj(2).NAXIS3 = NAXIS3 : StObj(2).XOffset = 1 : StObj(2).YOffset = 0
        StObj(3).NAXIS3 = NAXIS3 : StObj(3).XOffset = 1 : StObj(3).YOffset = 1

        'Start all threads
        For Each Slice As cStatMultiThread.cStatObjFixed In StObj
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
        Next Slice

        'Join all threads
        Do
            'System.Threading.Thread.Sleep(1)
            Dim AllDone As Boolean = True
            For Each Slice As cStatMultiThread.cStatObjFixed In StObj
                If Slice.Done = False Then
                    AllDone = False : Exit For
                End If
            Next Slice
            If AllDone Then Exit Do
        Loop Until 1 = 0

        'Collect all results
        ReDim Results(1, 1)
        Results(0, 0) = StObj(0)
        Results(0, 1) = StObj(1)
        Results(1, 0) = StObj(2)
        Results(1, 1) = StObj(3)

    End Sub

    '''<summary>Histogramm calculation itself - the histogram of one bayer channel is calculated.</summary>
    '''<remarks>Histogram is calculated via a the basic fixed-point dictionary as a vector would be 2^32 = 4 GByte entries long.</remarks>
    Private Sub HistoCalc(ByVal Arguments As Object)

        Dim StateObj As cStatMultiThread.cStatObjFixed = CType(Arguments, cStatMultiThread.cStatObjFixed)
        StateObj.Done = False

        'Count one bayer part
        StateObj.HistDataBayer = New Dictionary(Of ADUFixed, ADUCount)
        For IdxX As Integer = StateObj.XOffset To ImageData(StateObj.NAXIS3).Data.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData(StateObj.NAXIS3).Data.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                StateObj.HistDataBayer.AddTo(ImageData(StateObj.NAXIS3).Data(IdxX, IdxY), cStatMultiThread.OneUInt64)
            Next IdxY
        Next IdxX

        'Sort dictionary and return "done"
        StateObj.HistDataBayer = StateObj.HistDataBayer.SortDictionary
        StateObj.Done = True

    End Sub

End Class

'=================================================================================================================================
' Int32
'=================================================================================================================================

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
Public Class cStatMultiThread_Int32

    '''<summary>The real image data.</summary>
    Public ImageData(,) As Int32

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub RunHistoCalc(ByRef Results(,) As cStatMultiThread.cStatObjFixed)

        'Data are processed
        Dim StObj(3) As cStatMultiThread.cStatObjFixed
        For Idx As Integer = 0 To StObj.GetUpperBound(0)
            StObj(Idx) = New cStatMultiThread.cStatObjFixed
        Next Idx
        StObj(0).XOffset = 0 : StObj(0).YOffset = 0
        StObj(1).XOffset = 0 : StObj(1).YOffset = 1
        StObj(2).XOffset = 1 : StObj(2).YOffset = 0
        StObj(3).XOffset = 1 : StObj(3).YOffset = 1

        'Start all threads
        For Each Slice As cStatMultiThread.cStatObjFixed In StObj
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
        Next Slice

        'Join all threads
        Do
            'System.Threading.Thread.Sleep(1)
            Dim AllDone As Boolean = True
            For Each Slice As cStatMultiThread.cStatObjFixed In StObj
                If Slice.Done = False Then
                    AllDone = False : Exit For
                End If
            Next Slice
            If AllDone Then Exit Do
        Loop Until 1 = 0

        'Collect all results
        ReDim Results(1, 1)
        Results(0, 0) = StObj(0)
        Results(0, 1) = StObj(1)
        Results(1, 0) = StObj(2)
        Results(1, 1) = StObj(3)

    End Sub

    '''<summary>Histogramm calculation itself - the histogram of one bayer channel is calculated.</summary>
    Private Sub HistoCalc(ByVal Arguments As Object)

        Dim StateObj As cStatMultiThread.cStatObjFixed = CType(Arguments, cStatMultiThread.cStatObjFixed)
        StateObj.Done = False

        'Count one bayer part
        StateObj.HistDataBayer = New Dictionary(Of Int64, UInt64)
        For IdxX As Integer = StateObj.XOffset To ImageData.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                Dim PixelValue As Int32 = ImageData(IdxX, IdxY)
                If StateObj.HistDataBayer.ContainsKey(PixelValue) = False Then
                    StateObj.HistDataBayer.Add(PixelValue, cStatMultiThread.OneUInt64)
                Else
                    StateObj.HistDataBayer(PixelValue) += cStatMultiThread.OneUInt64
                End If
            Next IdxY
        Next IdxX

        'Sort dictionary and return "done"
        StateObj.HistDataBayer = StateObj.HistDataBayer.SortDictionary
        StateObj.Done = True

    End Sub

End Class

'=================================================================================================================================
' Float32
'=================================================================================================================================

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
Public Class cStatMultiThread_Float32

    Public Structure sImgData_Float32
        Public Data(,) As Single
        Public ReadOnly Property Length() As Long
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.LongLength
            End Get
        End Property
        Public ReadOnly Property NAXIS1() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(0) + 1
            End Get
        End Property
        Public ReadOnly Property NAXIS2() As Integer
            Get
                If IsNothing(Data) = True Then Return -1 Else Return Data.GetUpperBound(1) + 1
            End Get
        End Property
    End Structure

    '''<summary>The real image data.</summary>
    Public ImageData(3) As sImgData_Float32

    '''<summary>Object for each thread.</summary>
    Public Class cStateObj
        Friend XOffset As Integer = -1
        Friend YOffset As Integer = -1
        Friend HistDataBayer As New Dictionary(Of Single, UInt64)
        Friend Done As Boolean = False
    End Class

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub RunHistoCalc(ByRef Results(,) As cStateObj)

        'Data are processed
        Dim StObj(3) As cStateObj
        For Idx As Integer = 0 To StObj.GetUpperBound(0)
            StObj(Idx) = New cStateObj
        Next Idx
        StObj(0).XOffset = 0 : StObj(0).YOffset = 0
        StObj(1).XOffset = 0 : StObj(1).YOffset = 1
        StObj(2).XOffset = 1 : StObj(2).YOffset = 0
        StObj(3).XOffset = 1 : StObj(3).YOffset = 1

        'Start all threads
        For Each Slice As cStateObj In StObj
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
        Next Slice

        'Join all threads
        Do
            'System.Threading.Thread.Sleep(1)
            Dim AllDone As Boolean = True
            For Each Slice As cStateObj In StObj
                If Slice.Done = False Then
                    AllDone = False : Exit For
                End If
            Next Slice
            If AllDone Then Exit Do
        Loop Until 1 = 0

        'Collect all results
        ReDim Results(1, 1)
        Results(0, 0) = StObj(0)
        Results(0, 1) = StObj(1)
        Results(1, 0) = StObj(2)
        Results(1, 1) = StObj(3)

    End Sub

    '''<summary>Histogramm calculation itself - the histogram of one bayer channel is calculated.</summary>
    Private Sub HistoCalc(ByVal Arguments As Object)

        'TODO: Correct color processing

        Dim StateObj As cStateObj = CType(Arguments, cStateObj)
        StateObj.Done = False

        'Count one bayer part
        StateObj.HistDataBayer = New Dictionary(Of Single, UInt64)
        For IdxX As Integer = StateObj.XOffset To ImageData(0).Data.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData(0).Data.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                Dim PixelValue As Single = ImageData(0).Data(IdxX, IdxY)
                If StateObj.HistDataBayer.ContainsKey(PixelValue) = False Then
                    StateObj.HistDataBayer.Add(PixelValue, cStatMultiThread.OneUInt64)
                Else
                    StateObj.HistDataBayer(PixelValue) += cStatMultiThread.OneUInt64
                End If
            Next IdxY
        Next IdxX

        'Sort dictionary and return "done"
        StateObj.HistDataBayer = StateObj.HistDataBayer.SortDictionary
        StateObj.Done = True

    End Sub

End Class