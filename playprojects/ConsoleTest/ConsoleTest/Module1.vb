Option Explicit On
Option Strict On

Module Module1

    Const IPPPath As String = "C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2019.1.144\windows\redist\intel64_win\ipp"

    Sub Main()

        'Move and rotate an image
        'Example: http://hpc.ipp.ac.cn/wp-content/uploads/2015/12/documentation_2016/en/ipp/common/tutorials/ipp_blur_rotate/GUID-3B4A1095-6A5D-474C-A620-CDC265B18843.htm

        Dim X As New cIntelIPP(IPPPath)

        Dim coeffs(,) As Double = {}
        Dim Src As New cIntelIPP.sIppiSize(1000, 2000)
        Dim dst As New cIntelIPP.sIppiSize(2000, 4000)
        Dim DataType As cIntelIPP.eIppDataType = cIntelIPP.eIppDataType.ipp16u
        Dim Interpolation As cIntelIPP.eIppiInterpolationType = cIntelIPP.eIppiInterpolationType.ippCubic
        Dim Direction As cIntelIPP.eIppiWarpDirection = cIntelIPP.eIppiWarpDirection.ippWarpForward
        Dim BorderType As cIntelIPP.eIppiBorderType = cIntelIPP.eIppiBorderType.ippBorderConst

        Dim pSpecSize As Integer = -1
        Dim pInitBufSize As Integer = -1

        X.GetRotateTransform(10, 20, 30, coeffs)
        Dim Status As cIntelIPP.IppStatus = X.WarpAffineGetSize(Src, dst, DataType, coeffs, Interpolation, Direction, BorderType, pSpecSize, pInitBufSize)

    End Sub

    Sub TestImageCopy()

        Dim TextW As Integer = 2
        Dim X As New cIntelIPP(IPPPath)

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
