Module Module1

    Sub Main()

        Dim TextW As Integer = 2
        Dim X As New cIntelIPP("C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2019.1.144\windows\redist\intel64_win\ipp")

        Dim TestArray(,) As UInt16 = cIntelIPP.BuildTestData(11, 10)

        Console.WriteLine("INPUT DATA:")
        cIntelIPP.DisplayArray(TestArray, TextW)
        Console.WriteLine("---------------------------------------")
        cIntelIPP.DisplayArray(cIntelIPP.ShowHowInMemory(TestArray), TextW)
        Console.WriteLine("========================================")

        Dim NewArray(,) As UInt16 = {}
        X.Copy(TestArray, NewArray, 2, 3, 3, 5)

        Console.WriteLine("OUTPUT DATA:")
        cIntelIPP.DisplayArray(NewArray, TextW)
        Console.WriteLine("========================================")

        Console.WriteLine("DONE")
        Console.ReadKey()

    End Sub


End Module
