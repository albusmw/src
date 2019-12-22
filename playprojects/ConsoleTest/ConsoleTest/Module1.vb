Option Explicit On
Option Strict On

Module Module1

    'cImageMetaData.AddIPTCData("C:\TEMP\Astro Miesbach\Entwickelt\NGC7023 - Irisnebel.jpg")

    Const IPPPath As String = "C:\Program Files (x86)\IntelSWTools\compilers_and_libraries_2019.1.144\windows\redist\intel64_win\ipp"

    Public Sub Main()

        'Init table for conversion of UIn16 values to bytes to write
        Dim Int16UsignedToFITS As Int32 = 32768
        Dim UInt16Table As New Dictionary(Of UShort, Byte())
        For InDat As UInt16 = UInt16.MinValue To UInt16.MaxValue - 1
            UInt16Table.Add(InDat, GetBytesToStore(CType(InDat - Int16UsignedToFITS, Int16)))
        Next InDat
        UInt16Table.Add(UInt16.MaxValue, GetBytesToStore(CType(UInt16.MaxValue - Int16UsignedToFITS, Int16)))

        Dim IPPInput As New List(Of UInt16)
        For InDat As UInt16 = UInt16.MinValue To UInt16.MaxValue - 1
            IPPInput.Add(InDat)
        Next InDat
        IPPInput.Add(UInt16.MaxValue)

        Dim IPPVector As UInt16() = IPPInput.ToArray
        Dim Bytes((2 * IPPVector.Length) - 1) As Byte

        Dim X As New cIntelIPP(IPPPath)
        X.SwapBytes(IPPVector)
        X.XorC(IPPVector, &H80)

        For Idx As UInt16 = UInt16.MinValue To UInt16.MaxValue - 1
            Dim Correct As Byte() = UInt16Table(Idx)
            Dim FromIPP As Byte() = BitConverter.GetBytes(IPPVector(Idx))
            If Correct(0) <> FromIPP(0) Then
                MsgBox("ERROR as " & Idx.ToString.Trim, MsgBoxStyle.Exclamation)
                Exit For
            End If
            If Correct(1) <> FromIPP(1) Then
                MsgBox("ERROR as " & Idx.ToString.Trim, MsgBoxStyle.Exclamation)
                Exit For
            End If
        Next Idx

    End Sub

    Sub TestImageWarp()

        Dim Vector As Double() = {1, 2, 3, 4, 5, 6, 7, 8}

        'Move and rotate an image
        'Example: http://hpc.ipp.ac.cn/wp-content/uploads/2015/12/documentation_2016/en/ipp/common/tutorials/ipp_blur_rotate/GUID-3B4A1095-6A5D-474C-A620-CDC265B18843.htm

        Dim X As New cIntelIPP(IPPPath)

        Dim Minimum As Double = Double.NaN
        Dim Maximum As Double = Double.NaN
        X.MinMax(Vector, Minimum, Maximum)

        Dim coeffs(,) As Double = {}
        Dim Src As New cIntelIPP.IppiSize(1000, 2000)
        Dim dst As New cIntelIPP.IppiSize(1000, 2000)
        Dim DataType As cIntelIPP.IppDataType = cIntelIPP.IppDataType.ipp8u
        Dim Interpolation As cIntelIPP.IppiInterpolationType = cIntelIPP.IppiInterpolationType.ippLinear
        Dim Direction As cIntelIPP.IppiWarpDirection = cIntelIPP.IppiWarpDirection.ippWarpForward
        Dim BorderType As cIntelIPP.IppiBorderType = cIntelIPP.IppiBorderType.ippBorderConst

        Dim pSpecSize As Integer = -1
        Dim pInitBufSize As Integer = -1

        X.GetRotateTransform(100, 20, 30, coeffs)
        coeffs = {{1, 2, 3}, {4, 5, 6}}
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

    Private Function GetBytesToStore(ByVal Value As Int16) As Byte()
        Dim RetVal As Byte() = BitConverter.GetBytes(Value)
        Return New Byte() {RetVal(1), RetVal(0)}
    End Function

End Module
