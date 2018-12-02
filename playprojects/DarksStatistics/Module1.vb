Module Module1

    Sub Main()

        Dim Sum(,) As Double = {}
        Dim SumSum(,) As Double = {}
        Dim ImageCount As Integer = 0

        For Each File As String In System.IO.Directory.GetFiles("C:\Users\Martin Weiss\Downloads\M51\1000Da\1000Da\Bias", "*.cr2")
            Dim FileNameOnly As String = System.IO.Path.GetFileName(File)
            Console.WriteLine(FileNameOnly & " ...")
            Dim Output() As String = {}
            Dim GeneratedFile As String = cImageFileFormatReader.UseDCRaw("C:\Users\Martin Weiss\Dropbox\Bin\dcraw-9.27-ms-64-bit.exe", File, Output)
            Console.WriteLine("  >> " & GeneratedFile)
            Dim ImageData(,) As Double = {}
            cImageFileFormatReader.LoadPortableAnyMapNew(GeneratedFile, ImageData, 0)
            ImageCount += 1
            If Sum.LongLength = 0 Then
                ReDim Sum(ImageData.GetUpperBound(0), ImageData.GetUpperBound(1))
                ReDim SumSum(ImageData.GetUpperBound(0), ImageData.GetUpperBound(1))
                Console.WriteLine("  >> Statistics will run on <" & (ImageData.GetUpperBound(0) + 1).ToString.Trim & "x" & (ImageData.GetUpperBound(1) + 1).ToString.Trim & ">")
            End If
            Console.WriteLine("  >> Current image is <" & (ImageData.GetUpperBound(0) + 1).ToString.Trim & "x" & (ImageData.GetUpperBound(1) + 1).ToString.Trim & ">")
            If ImageData.GetUpperBound(0) = Sum.GetUpperBound(0) And ImageData.GetUpperBound(1) = Sum.GetUpperBound(1) Then
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(0)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(1)
                        Sum(Idx1, Idx2) += ImageData(Idx1, Idx2)
                        SumSum(Idx1, Idx2) += ImageData(Idx1, Idx2) * ImageData(Idx1, Idx2)
                    Next Idx2
                Next Idx1
            Else
                If ImageData.GetUpperBound(0) = Sum.GetUpperBound(1) And ImageData.GetUpperBound(1) = Sum.GetUpperBound(0) Then
                    For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                        For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                            Sum(Idx1, Idx2) += ImageData(Idx2, Idx1)
                            SumSum(Idx1, Idx2) += ImageData(Idx2, Idx1) * ImageData(Idx2, Idx1)
                        Next Idx2
                    Next Idx1
                End If
            End If
            System.IO.File.Delete(GeneratedFile)
        Next File

        Dim Mean(Sum.GetUpperBound(0), Sum.GetUpperBound(1)) As Double
        Dim Sigma(Sum.GetUpperBound(0), Sum.GetUpperBound(1)) As Double
        For Idx1 As Integer = 0 To Sum.GetUpperBound(0)
            For Idx2 As Integer = 0 To Sum.GetUpperBound(1)
                Sigma(Idx1, Idx2) = System.Math.Sqrt((SumSum(Idx1, Idx2) - ((Sum(Idx1, Idx2) * Sum(Idx1, Idx2)) / ImageCount)) / (ImageCount - 1))
                Mean(Idx1, Idx2) = Sum(Idx1, Idx2) / ImageCount
            Next Idx2
        Next Idx1

        Dim FitsToWrite As String = "C:\Users\Martin Weiss\Downloads\M51\1000Da\1000Da\Bias\Sigma.fits"
        cFITSWriter.Write(FitsToWrite, Sigma, cFITSWriter.eBitPix.Double)
        cFITSWriter.Write(FitsToWrite.Replace("Sigma.", "Mean."), Mean, cFITSWriter.eBitPix.Double)
        Process.Start(FitsToWrite)

        Console.WriteLine("=========================================================")
        Console.WriteLine("DONE")
        Console.ReadKey()

    End Sub

End Module
