Option Explicit On
Option Strict On

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
Public Class cStatMultiThread(Of T)

    Private Const OneUInt32 As UInt32 = CType(1, UInt32)

    Public Data(,) As T

    '''<summary>Object for each thread.</summary>
    Public Class cStateObj(Of DType)
        Friend StartIdx As Integer = -1
        Friend StopIdx As Integer = -1
        Friend HistDataBayer(,) As Dictionary(Of DType, UInt32)
        Friend Done As Boolean = False
        '''<summary>By default we use a 2x2 bayer histogram calculation.</summary>
        Public Sub New()
            ReDim HistDataBayer(1, 1)
            For Idx1 As Integer = 0 To HistDataBayer.GetUpperBound(0)
                For Idx2 As Integer = 0 To HistDataBayer.GetUpperBound(0)
                    HistDataBayer(Idx1, Idx2) = New Dictionary(Of DType, UInt32)
                Next Idx2
            Next Idx1
        End Sub
    End Class

    '''<summary>Perform a calculation with the given number of threads.</summary>
    Public Sub Calculate(ByVal ThreadCount As Integer, ByRef Results As cStateObj(Of T))

        Dim SliceSize As Integer = 2 * ((Data.GetUpperBound(0) \ ThreadCount) \ 2)
        Dim StObj(ThreadCount - 1) As cStateObj(Of T)
        For Idx As Integer = 0 To StObj.GetUpperBound(0)
            StObj(Idx) = New cStateObj(Of T)
        Next Idx
        StObj(0).StartIdx = 0
        StObj(0).StopIdx = SliceSize
        StObj(StObj.GetUpperBound(0)).StopIdx = Data.GetUpperBound(0) - 1
        For Idx As Integer = 1 To StObj.GetUpperBound(0) - 1
            StObj(Idx).StartIdx = StObj(Idx - 1).StopIdx + 2
            StObj(Idx).StopIdx = StObj(Idx).StartIdx + SliceSize
        Next Idx
        StObj(StObj.GetUpperBound(0)).StartIdx = StObj(StObj.GetUpperBound(0) - 1).StopIdx + 2

        'Start all threads
        For Each Slice As cStateObj(Of T) In StObj
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
        Next Slice

        'Join all threads
        Do
            System.Threading.Thread.Sleep(1)
            Dim AllDone As Boolean = True
            For Each Slice As cStateObj(Of T) In StObj
                If Slice.Done = False Then
                    AllDone = False : Exit For
                End If
            Next Slice
            If AllDone Then Exit Do
        Loop Until 1 = 0

        'Collect all results
        Results = New cStateObj(Of T)
        'Combine bayer results of each thread
        For Each Slice As cStateObj(Of T) In StObj
            'Combine the bayer matrix histogram
            For BayerX As Integer = 0 To 1
                For BayerY As Integer = 0 To 1
                    For Each PixelValue As T In Slice.HistDataBayer(BayerX, BayerY).Keys
                        Dim BinCount As UInt32 = Slice.HistDataBayer(BayerX, BayerY)(PixelValue)
                        If Results.HistDataBayer(BayerX, BayerY).ContainsKey(PixelValue) = False Then
                            Results.HistDataBayer(BayerX, BayerY).Add(PixelValue, BinCount)
                        Else
                            Results.HistDataBayer(BayerX, BayerY)(PixelValue) += BinCount
                        End If
                    Next PixelValue
                Next BayerY
            Next BayerX
        Next Slice

        'Post-calculation
        Results.HistDataBayer(0, 0) = cGenerics.SortDictionary(Results.HistDataBayer(0, 0))
        Results.HistDataBayer(0, 1) = cGenerics.SortDictionary(Results.HistDataBayer(0, 1))
        Results.HistDataBayer(1, 0) = cGenerics.SortDictionary(Results.HistDataBayer(1, 0))
        Results.HistDataBayer(1, 1) = cGenerics.SortDictionary(Results.HistDataBayer(1, 1))

    End Sub

    '''<summary>Histogramm calculation itself.</summary>
    Private Sub HistoCalc(ByVal Arguments As Object)

        Dim StateObj As cStateObj(Of T) = CType(Arguments, cStateObj(Of T))
        StateObj.Done = False

        For IdxX As Integer = StateObj.StartIdx To StateObj.StopIdx Step 2
            For IdxY As Integer = 0 To Data.GetUpperBound(1) - 1 Step 2
                'Calculate a separat histogram for each bayer matrix element
                For BayerX As Integer = 0 To 1
                    For BayerY As Integer = 0 To 1
                        Dim PixelValue As T = Data(IdxX + BayerX, IdxY + BayerY)
                        If StateObj.HistDataBayer(BayerX, BayerY).ContainsKey(PixelValue) = False Then
                            StateObj.HistDataBayer(BayerX, BayerY).Add(PixelValue, OneUInt32)
                        Else
                            StateObj.HistDataBayer(BayerX, BayerY)(PixelValue) += OneUInt32
                        End If
                    Next BayerY
                Next BayerX
            Next IdxY
        Next IdxX

        StateObj.Done = True

    End Sub


End Class