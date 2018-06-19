Option Explicit On
Option Strict On

'''<summary>Class to calculate 2D matrix statistics multi-threaded.</summary>
Public Class cStatMultiThread

  Public Data(,) As Double

  '''<summary>Object for each thread.</summary>
  Public Class cStateObj
    Friend StartIdx As Integer
    Friend StopIdx As Integer
    Friend HistData As New Dictionary(Of Double, Integer)
    Friend HistDataBayer(,) As Dictionary(Of Double, Integer)
    Friend Max As Double = Double.MinValue
    Friend Min As Double = Double.MaxValue
    Friend Sum As Double = 0.0
    Friend Done As Boolean = False
    Public Sub New()
      ReDim HistDataBayer(1, 1)
      For Idx1 As Integer = 0 To HistDataBayer.GetUpperBound(0)
        For Idx2 As Integer = 0 To HistDataBayer.GetUpperBound(0)
          HistDataBayer(Idx1, Idx2) = New Dictionary(Of Double, Integer)
        Next Idx2
      Next Idx1
    End Sub
  End Class

  '''<summary>Perform a calculation with the given number of threads.</summary>
  Public Sub Calculate(ByVal ThreadCount As Integer, ByRef Results As cStateObj)

    Dim SliceSize As Integer = Data.GetUpperBound(0) \ ThreadCount
    If SliceSize Mod 2 <> 0 And ThreadCount > 1 Then SliceSize += 1     'ensure an even slice size ...
    Dim StObj(ThreadCount - 1) As cStateObj
    For Idx As Integer = 0 To StObj.GetUpperBound(0)
      StObj(Idx) = New cStateObj()
      Select Case Idx
        Case 0
          'first slice will start from 0 on
          StObj(Idx).StartIdx = 0
          StObj(Idx).StopIdx = SliceSize
        Case StObj.GetUpperBound(0)
          'last slice will end at the end
          StObj(Idx).StartIdx = StObj(Idx - 1).StopIdx + 1
          StObj(Idx).StopIdx = Data.GetUpperBound(0)
        Case Else
          'other slices will go in between
          StObj(Idx).StartIdx = StObj(Idx - 1).StopIdx + 1
          StObj(Idx).StopIdx = StObj(Idx).StartIdx + SliceSize
      End Select
    Next Idx

    'Start all threads
    For Each Slice As cStateObj In StObj
      System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf HistoCalc), Slice)
    Next Slice

    'Join all threads
    Do
      System.Threading.Thread.Sleep(5)
      Dim AllDone As Boolean = True
      For Each Slice As cStateObj In StObj
        If Slice.Done = False Then
          AllDone = False : Exit For
        End If
      Next Slice
      If AllDone Then Exit Do
    Loop Until 1 = 0

    'Collect all results
    Results = New cStateObj
    For Each Slice As cStateObj In StObj
      'Combine the bayer matrix histogram
      For BIdx1 As Integer = 0 To 1
        For BIdx2 As Integer = 0 To 1
          For Each PixelValue As Double In Slice.HistDataBayer(BIdx1, BIdx2).Keys
            Dim BinCount As Integer = Slice.HistDataBayer(BIdx1, BIdx2)(PixelValue)
            'Combine bayer results of each thread
            If Results.HistDataBayer(BIdx1, BIdx2).ContainsKey(PixelValue) = False Then
              Results.HistDataBayer(BIdx1, BIdx2).Add(PixelValue, BinCount)
            Else
              Results.HistDataBayer(BIdx1, BIdx2)(PixelValue) += BinCount
            End If
            'Create overall histogramm
            If Results.HistData.ContainsKey(PixelValue) = False Then
              Results.HistData.Add(PixelValue, BinCount)
            Else
              Results.HistData(PixelValue) += BinCount
            End If
          Next PixelValue
        Next BIdx2
      Next BIdx1
      Results.Sum += Slice.Sum
    Next Slice

    'Post-calculation
    Results.HistData = cGenerics.SortDictionary(Results.HistData)
    Results.HistDataBayer(0, 0) = cGenerics.SortDictionary(Results.HistDataBayer(0, 0))
    Results.HistDataBayer(0, 1) = cGenerics.SortDictionary(Results.HistDataBayer(0, 1))
    Results.HistDataBayer(1, 0) = cGenerics.SortDictionary(Results.HistDataBayer(1, 0))
    Results.HistDataBayer(1, 1) = cGenerics.SortDictionary(Results.HistDataBayer(1, 1))

    Results.Min = cGenerics.GetDictionaryKeyElement(Results.HistData, 0)
    Results.Max = cGenerics.GetDictionaryKeyElement(Results.HistData, Results.HistData.Count - 1)

  End Sub

  '''<summary>Histogramm calculation itself.</summary>
  Public Sub HistoCalc(ByVal Arguments As Object)

    Dim StateObj As cStateObj = CType(Arguments, cStateObj)

    For Idx1 As Integer = StateObj.StartIdx To StateObj.StopIdx - 1 Step 2
      For Idx2 As Integer = 0 To Data.GetUpperBound(1) - 1 Step 2
        'Calculate a separat histogram for each bayer matrix element
        For BIdx1 As Integer = 0 To 1
          For BIdx2 As Integer = 0 To 1
            Dim PixelValue As Double = Data(Idx1 + BIdx1, Idx2 + BIdx2)
            If StateObj.HistDataBayer(BIdx1, BIdx2).ContainsKey(PixelValue) = False Then
              StateObj.HistDataBayer(BIdx1, BIdx2).Add(PixelValue, 1)
            Else
              StateObj.HistDataBayer(BIdx1, BIdx2)(PixelValue) += 1
            End If
            StateObj.Sum += PixelValue
          Next BIdx2
        Next BIdx1
      Next Idx2
    Next Idx1

    StateObj.Done = True

  End Sub


End Class