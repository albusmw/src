Option Explicit On
Option Strict On

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
'''<remarks>Calculation is done by buidling a vector with all possible entries (only 2^16 length).</remarks>
Public Class cStatMultiThread_UInt16

    Private Const OneUInt64 As UInt64 = CType(1, UInt64)

    '''<summary>The real image data.</summary>
    Public ImageData(,) As UInt16

    '''<summary>Object for each thread.</summary>
    Public Class cStateObj
        Friend XOffset As Integer = -1
        Friend YOffset As Integer = -1
        Friend HistDataBayer As New Dictionary(Of Int64, UInt64)
        Friend Done As Boolean = False
    End Class

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub Calculate(ByRef Results(,) As cStateObj)

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

        Dim StateObj As cStateObj = CType(Arguments, cStateObj)
        StateObj.Done = False

        'Init count object with 0
        Dim HistCount(UInt16.MaxValue) As UInt64
        For Idx As Integer = 0 To HistCount.GetUpperBound(0)
            HistCount(Idx) = 0
        Next Idx

        'Count one bayer part
        Dim XOffsets As New List(Of Integer)
        For IdxX As Integer = StateObj.XOffset To ImageData.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            XOffsets.Add(IdxX)
        Next IdxX
        Threading.Tasks.Parallel.ForEach(XOffsets, Sub(IdxX)
                                                       For IdxY As Integer = StateObj.YOffset To ImageData.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                                                           HistCount(ImageData(IdxX, IdxY)) += OneUInt64
                                                       Next IdxY
                                                   End Sub)

        'Form return value
        StateObj.HistDataBayer = New Dictionary(Of Int64, UInt64)
        For Idx As UInt16 = 0 To CUShort(HistCount.GetUpperBound(0))
            If HistCount(Idx) > 0 Then StateObj.HistDataBayer.Add(Idx, HistCount(Idx))
            If Idx = HistCount.GetUpperBound(0) Then Exit For
        Next Idx

        StateObj.Done = True

    End Sub

End Class

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
'''<remarks>Calculation as for UInt16 is not possible as a vector with all entries would be 2^32 entries long.</remarks>
Public Class cStatMultiThread_UInt32

    Private Const OneUInt64 As UInt64 = CType(1, UInt64)

    '''<summary>The real image data.</summary>
    Public ImageData(,) As UInt32

    '''<summary>Object for each thread.</summary>
    Public Class cStateObj
        Friend XOffset As Integer = -1
        Friend YOffset As Integer = -1
        Friend HistDataBayer As New Dictionary(Of Int64, UInt64)
        Friend Done As Boolean = False
    End Class

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub Calculate(ByRef Results(,) As cStateObj)

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

        Dim StateObj As cStateObj = CType(Arguments, cStateObj)
        StateObj.Done = False

        'Count one bayer part
        StateObj.HistDataBayer = New Dictionary(Of Int64, UInt64)
        For IdxX As Integer = StateObj.XOffset To ImageData.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                Dim PixelValue As UInt32 = ImageData(IdxX, IdxY)
                If StateObj.HistDataBayer.ContainsKey(PixelValue) = False Then
                    StateObj.HistDataBayer.Add(PixelValue, OneUInt64)
                Else
                    StateObj.HistDataBayer(PixelValue) += OneUInt64
                End If
            Next IdxY
        Next IdxX

        StateObj.Done = True

    End Sub

End Class

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
Public Class cStatMultiThread_Int32

    Private Const OneInt64 As UInt64 = CType(1, Int64)

    '''<summary>The real image data.</summary>
    Public ImageData(,) As Int32

    '''<summary>Object for each thread.</summary>
    Public Class cStateObj
        Friend XOffset As Integer = -1
        Friend YOffset As Integer = -1
        Friend HistDataBayer As New Dictionary(Of Int64, UInt64)
        Friend Done As Boolean = False
    End Class

    '''<summary>Perform a calculation with 4 threads, one for each bayer channel.</summary>
    Public Sub Calculate(ByRef Results(,) As cStateObj)

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

        Dim StateObj As cStateObj = CType(Arguments, cStateObj)
        StateObj.Done = False

        'Count one bayer part
        StateObj.HistDataBayer = New Dictionary(Of Int64, UInt64)
        For IdxX As Integer = StateObj.XOffset To ImageData.GetUpperBound(0) - 1 + StateObj.XOffset Step 2
            For IdxY As Integer = StateObj.YOffset To ImageData.GetUpperBound(1) - 1 + StateObj.YOffset Step 2
                Dim PixelValue As Int32 = ImageData(IdxX, IdxY)
                If StateObj.HistDataBayer.ContainsKey(PixelValue) = False Then
                    StateObj.HistDataBayer.Add(PixelValue, OneInt64)
                Else
                    StateObj.HistDataBayer(PixelValue) += OneInt64
                End If
            Next IdxY
        Next IdxX

        'Sort dictionary and return "done"
        StateObj.HistDataBayer = StateObj.HistDataBayer.SortDictionary
        StateObj.Done = True

    End Sub

End Class