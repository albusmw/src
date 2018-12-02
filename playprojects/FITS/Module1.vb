Module Module1

    Sub Main()

        Dim X As New cFITSReader
        Dim ImageData(,) As Double

        X.ReadIn("C:\Users\Martin Weiss\Dropbox\Astro\von  Hans\Flats_0,005-001flat.fit", ImageData)
        Dim FITSFolder As String = "C:\Users\Martin Weiss\Dropbox\Astro\von  Hans"

        Dim XSum(ImageData.GetUpperBound(0)) As Double
        Dim YSum(ImageData.GetUpperBound(1)) As Double
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(0)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(1)
                XSum(Idx1) += ImageData(Idx1, Idx2)
                YSum(Idx2) += ImageData(Idx1, Idx2)
            Next Idx2
        Next Idx1

        For Idx As Integer = 0 To XSum.GetUpperBound(0)
            XSum(Idx) /= (ImageData.GetUpperBound(1) + 1)
        Next Idx
        For Idx As Integer = 0 To YSum.GetUpperBound(0)
            YSum(Idx) /= (ImageData.GetUpperBound(0) + 1)
        Next Idx

        Dim Histo As New List(Of String)
        For Idx1 As Integer = 0 To Math.Max(XSum.GetUpperBound(0), YSum.GetUpperBound(0))
            Dim CSVLine As New List(Of String)
            CSVLine.Add(Idx1.ToString.Trim)
            If Idx1 <= XSum.GetUpperBound(0) Then CSVLine.Add(XSum(Idx1).ToString.Trim)
            If Idx1 <= YSum.GetUpperBound(0) Then CSVLine.Add(YSum(Idx1).ToString.Trim)
            Histo.Add(Join(CSVLine.ToArray, ";"))
        Next Idx1
        System.IO.File.WriteAllLines(FITSFolder & "\Histo1.csv", Histo.ToArray)
        Process.Start(FITSFolder & "\Histo1.csv")

    End Sub

End Module
