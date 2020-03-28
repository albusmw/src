Option Explicit On
Option Strict On

'TODO:
' - Implement UInt16 and UInt32 saving with BZERO usage

'''<summary>Class to write 2-dimensional arrays as FITS data.</summary>
Public Class cFITSWriter

    '''<summary>Value that is stored if the passed data could not be stored as byte value..</summary>
    Private Const ByteValueInvalid As Byte = 0
    '''<summary>Value that is stored if the passed data could not be stored as Int16 value..</summary>
    Private Const Int16ValueInvalid As Int16 = 0
    '''<summary>Value that is stored if the passed data could not be stored as Int32 value..</summary>
    Private Const Int32ValueInvalid As Int32 = 0
    '''<summary>Value that is stored if the passed data could not be stored as Single value..</summary>
    Private Const SingleValueInvalid As Single = Single.NaN

    Private Shared UInt16Table As Collections.Generic.Dictionary(Of UInt16, Byte())

    '''<summary>Path to ipps.dll and ippvm.dll - if not set IPP will not be used.</summary>
    Public Shared Property IPPPath As String = String.Empty

    Public Shared Property UseIPPForWriting As Boolean = False

    Private Const BZeroNotUsed As Double = 0.0
    Private Const BScaleNotUsed As Double = 1.0
    Private Const Int16UsignedToFITS As Int32 = 32768
    Private Const Int32UsignedToFITS As Int64 = 2147483648
    Private Const Int64UsignedToFITS As Int64 = 9223372036854775807     'must be +1 but can not be displayed ...

    Public Class FITSWriterException : Inherits Exception
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    '- Header data that are directly dependant on the concent are stored automatically
    '- Additional data are passed as list of vectors

    Public Enum eBitPix
        [Byte] = 8
        [Int16] = 16
        [Int32] = 32
        [Single] = -32
        [Double] = -64
    End Enum

    '''<summary>Length of one header element.</summary>
    Public Shared Property HeaderElementLength As Integer = 80
    '''<summary>Length of a header block - FITS files may contain an integer size of header blocks.</summary>
    Public Shared Property HeaderBlockSize As Integer = 2880
    '''<summary>Length of the keyword entry.</summary>
    Public Shared Property KeywordLength As Integer = 8
    '''<summary>Length of the value entry.</summary>
    Public Shared Property ValueLength As Integer = 18
    '''<summary>Number of header elements per header block.</summary>
    Public Shared ReadOnly HeaderElements As Integer = HeaderBlockSize \ HeaderElementLength

    '================================================================================================
    ' Byte
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Byte, ByVal BitPix As eBitPix) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, 1.0, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Byte, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, CustomHeaderElements)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Byte, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Byte
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix8(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Int32 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '================================================================================================
    ' Int16
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int16, ByVal BitPix As eBitPix) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int16, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, CustomHeaderElements)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int16, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Int16
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Int32 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '================================================================================================
    ' UInt16
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt16, ByVal BitPix As eBitPix) As Integer
        Return Write(FileName, ImageData, BitPix, Int16UsignedToFITS, BScaleNotUsed, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt16, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return Write(FileName, ImageData, BitPix, Int16UsignedToFITS, BScaleNotUsed, CustomHeaderElements)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt16, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Init table for conversion of UIn16 values to bytes to write
        If IsNothing(UInt16Table) = True Then
            UInt16Table = New Collections.Generic.Dictionary(Of UShort, Byte())
            For InDat As UInt16 = UInt16.MinValue To UInt16.MaxValue - 1
                UInt16Table.Add(InDat, GetBytes_BitPix16(CType(InDat - Int16UsignedToFITS, Int16)))
            Next InDat
            UInt16Table.Add(UInt16.MaxValue, GetBytes_BitPix16(CType(UInt16.MaxValue - Int16UsignedToFITS, Int16)))
        End If

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Int16
                If BZero = Int16UsignedToFITS And BScale = BScaleNotUsed Then
                    'Write "as is" without any additional calculation (as there is no scaling ...); the only scaling needed is to subtrace 32768 in order to get a Int16 for the UInt16 ...
                    'We write the data blockwise to speed up writing ...
                    If UseIPPForWriting Then
                        Dim IntelIPP As New cIntelIPP(IPPPath)
                        Dim IPPStatus As New Collections.Generic.List(Of cIntelIPP.IppStatus)
                        Dim BytesToWrite((ImageData.Length * 2) - 1) As Byte
                        IPPStatus.Add(IntelIPP.SwapBytes(ImageData))
                        IPPStatus.Add(IntelIPP.XorC(ImageData, &H80))
                        IPPStatus.Add(IntelIPP.Transpose(ImageData, BytesToWrite))
                        BytesOut.Write(BytesToWrite)
                    Else
                        Dim Block(((ImageData.GetUpperBound(0) + 1) * 2) - 1) As Byte
                        Dim BlockPtr As Integer = 0
                        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                                Dim Val(1) As Byte : Val = UInt16Table(ImageData(Idx2, Idx1))
                                Block(BlockPtr) = Val(0) : Block(BlockPtr + 1) = Val(1)
                                BlockPtr += 2
                            Next Idx2
                            BytesOut.Write(Block) : BlockPtr = 0
                        Next Idx1
                    End If
                Else
                    'Write with scaling and offset taken into account
                    For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                        For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                            BytesOut.Write(GetBytes_BitPix16(ImageData(Idx2, Idx1), A, B, RetVal))
                        Next Idx2
                    Next Idx1
                End If
            Case Else
                Throw New FITSWriterException("Conversion from Int16 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '================================================================================================
    ' UInt32
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt32, ByVal BitPix As eBitPix) As Integer
        Return Write(FileName, ImageData, BitPix, Int32UsignedToFITS, BScaleNotUsed, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt32, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return Write(FileName, ImageData, BitPix, Int32UsignedToFITS, BScaleNotUsed, CustomHeaderElements)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As UInt32, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim BitPerPixel As Integer = 4
        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Dim UseIPP As Boolean = False
        Select Case BitPix
            Case eBitPix.Int32
                If BZero = Int32UsignedToFITS And BScale = BScaleNotUsed Then
                    'Write content as-is
                    'Write "as is" without any additional calculation (as there is no scaling ...); the only scaling needed is to subtrace Int64.MaxValue in order to get a Int16 for the UInt16 ...
                    'We write the data blockwise to speed up writing ...
                    If UseIPP = True Then
                        Dim IntelIPP As New cIntelIPP(IPPPath)
                        Dim IPPStatus As New Collections.Generic.List(Of cIntelIPP.IppStatus)
                        Dim BytesToWrite((ImageData.Length * 2) - 1) As Byte
                        Dim UnsignedXOR As UInt32 = BitConverter.ToUInt32(New Byte() {&H80, 0, 0, 0}, 0)
                        'IPPStatus.Add(IntelIPP.XorC(ImageData, UnsignedXOR))
                        'IPPStatus.Add(IntelIPP.SwapBytes(ImageData))
                        'IPPStatus.Add(IntelIPP.Transpose(ImageData, BytesToWrite))
                        BytesOut.Write(BytesToWrite)
                    Else
                        Dim Block(((ImageData.GetUpperBound(0) + 1) * BitPerPixel) - 1) As Byte
                        Dim BlockPtr As Integer = 0
                        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                                Dim Val(3) As Byte : Val = BitConverter.GetBytes(CType(ImageData(Idx2, Idx1) + Int32UsignedToFITS, UInt32))
                                Block(BlockPtr) = Val(3) : Block(BlockPtr + 1) = Val(2) : Block(BlockPtr + 2) = Val(1) : Block(BlockPtr + 3) = Val(0)
                                BlockPtr += BitPerPixel
                            Next Idx2
                            BytesOut.Write(Block) : BlockPtr = 0
                        Next Idx1
                    End If

                Else
                    'Write with scaling and offset taken into account
                    For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                        For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                            BytesOut.Write(GetBytes_BitPix32(ImageData(Idx2, Idx1), A, B, RetVal))
                        Next Idx2
                    Next Idx1
                End If
            Case Else
                Throw New FITSWriterException("Conversion from Int32 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '================================================================================================
    ' Int32
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int32, ByVal BitPix As eBitPix) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int32, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, CustomHeaderElements)
    End Function

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Int32, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Int16
                '1-to-1 copy
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Int32
                '1-to-1 copy
                If BZero = 0.0 And BScale = 1.0 Then
                    'No scaling ...
                    For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                        For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                            Dim DataToWrite As Byte() = BitConverter.GetBytes(ImageData(Idx2, Idx1))
                            BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                        Next Idx2
                    Next Idx1
                Else
                    For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                        For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                            BytesOut.Write(GetBytes_BitPix32(ImageData(Idx2, Idx1), A, B, RetVal))
                        Next Idx2
                    Next Idx1
                End If
            Case eBitPix.Single
                'fixed to floating point
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CSng(ImageData(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case eBitPix.Double
                'fixed to floating point
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CDbl(ImageData(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Int32 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '''<summary>Write the passed ImageData matrix to a color FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function WriteRGB(ByVal FileName As String, ByRef ImageDataR(,) As Int32, ByRef ImageDataG(,) As Int32, ByRef ImageDataB(,) As Int32, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer
        Return WriteRGB(FileName, ImageDataR, ImageDataG, ImageDataB, BitPix, BZeroNotUsed, BScaleNotUsed, Nothing)
    End Function

    '''<summary>Write the passed ImageData matrix to a color FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Function WriteRGB(ByVal FileName As String, ByRef ImageDataR(,) As Int32, ByRef ImageDataG(,) As Int32, ByRef ImageDataB(,) As Int32, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "3"))
        Header.Add(FormatHeader("NAXIS1", ImageDataR.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageDataR.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("NAXIS3", "3"))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Byte
                Dim ValueToStore As Byte = 0
                'R
                For Idx1 As Integer = 0 To ImageDataR.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataR.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix8(ImageDataR(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'G
                For Idx1 As Integer = 0 To ImageDataG.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataG.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix8(ImageDataG(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'B
                For Idx1 As Integer = 0 To ImageDataB.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataB.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix8(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Int16
                'R
                For Idx1 As Integer = 0 To ImageDataR.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataR.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'G
                For Idx1 As Integer = 0 To ImageDataG.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataG.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'B
                For Idx1 As Integer = 0 To ImageDataB.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataB.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Int32
                'R
                For Idx1 As Integer = 0 To ImageDataR.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataR.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix32(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'G
                For Idx1 As Integer = 0 To ImageDataG.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataG.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix32(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
                'B
                For Idx1 As Integer = 0 To ImageDataB.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataB.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix32(ImageDataB(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Int32 to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '================================================================================================
    ' Single
    '================================================================================================

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Sub Write(ByVal FileName As String, ByRef ImageData(,) As Single, ByVal BitPix As eBitPix)
        Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, Nothing)
    End Sub

    Public Shared Sub Write(ByVal FileName As String, ByRef ImageData(,) As Single, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String()))

        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        Select Case BitPix
            Case eBitPix.Single
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ImageData(Idx2, Idx1))
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case eBitPix.Double
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CDbl(ImageData(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Single to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '''<summary>Write the passed ImageData matrix to a color FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Sub Write(ByVal FileName As String, ByRef ImageDataR(,) As Single, ByRef ImageDataG(,) As Single, ByRef ImageDataB(,) As Single, ByVal BZero As Single, ByVal BScale As Single, ByVal BitPix As eBitPix, ByVal CustomHeaderElements As Collections.Generic.List(Of String()))

        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "3"))
        Header.Add(FormatHeader("NAXIS1", ImageDataR.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageDataR.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("NAXIS3", "3"))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Calculate scaler
        Dim A As Single = 1 / BScale
        Dim B As Single = -(BZero / BScale)

        'Write content
        Select Case BitPix
            Case eBitPix.Single
                '1-to-1 copy
                For Idx1 As Integer = 0 To ImageDataR.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataR.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ImageDataR(Idx2, Idx1))
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
                For Idx1 As Integer = 0 To ImageDataG.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataG.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ImageDataG(Idx2, Idx1))
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
                For Idx1 As Integer = 0 To ImageDataB.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataB.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ImageDataB(Idx2, Idx1))
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case eBitPix.Double
                '1-to-1 copy
                For Idx1 As Integer = 0 To ImageDataR.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataR.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CDbl(ImageDataR(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
                For Idx1 As Integer = 0 To ImageDataG.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataG.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CDbl(ImageDataG(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
                For Idx1 As Integer = 0 To ImageDataB.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageDataB.GetUpperBound(0)
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(CDbl(ImageDataB(Idx2, Idx1)))
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Single to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '================================================================================================
    ' Double
    '================================================================================================


    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    '''<param name="FileName">File to generate.</param>
    '''<param name="ImageData">Data to write.</param>
    '''<param name="BitPix">Bit per pixel.</param>
    Public Shared Sub Write(ByVal FileName As String, ByRef ImageData(,) As Double, ByVal BitPix As eBitPix)
        Write(FileName, ImageData, BitPix, BZeroNotUsed, BScaleNotUsed, Nothing)
    End Sub

    '''<summary>Write the passed ImageData matrix to a FITS file.</summary>
    ''' <param name="FileName">File name to generate.</param>
    ''' <param name="ImageData">Image data to store.</param>
    ''' <param name="BitPix">Bit-per-pixel according to FITS standard.</param>
    ''' <param name="BZero">BZero   of the formular RealValue = BZero + (BScale * StoredValue).</param>
    ''' <param name="BScale">BScale of the formular RealValue = BZero + (BScale * StoredValue).</param>
    ''' <param name="CustomHeaderElements">Custom FITS elements to store.</param>
    ''' <returns>Number of values that could NOT be stored.</returns>
    Public Shared Function Write(ByVal FileName As String, ByRef ImageData(,) As Double, ByVal BitPix As eBitPix, ByVal BZero As Double, ByVal BScale As Double, ByVal CustomHeaderElements As Collections.Generic.List(Of String())) As Integer

        Dim RetVal As Integer = 0
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", BZero.ValRegIndep.Replace(",", ".")))
        Header.Add(FormatHeader("BSCALE", BScale.ValRegIndep.Replace(",", ".")))
        AddCustomHeaders(Header, CustomHeaderElements)

        'Calculate scaler
        Dim A As Double = 1 / BScale
        Dim B As Double = -(BZero / BScale)

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        Select Case BitPix
            Case eBitPix.Byte
                Dim ValueToStore As Byte = 0
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix8(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Int16
                Dim ValueToStore As Int16 = 0
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix16(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Int32
                Dim ValueToStore As Int32 = 0
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        BytesOut.Write(GetBytes_BitPix32(ImageData(Idx2, Idx1), A, B, RetVal))
                    Next Idx2
                Next Idx1
            Case eBitPix.Single
                Dim ValueToStore As Single = 0
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim Scaled As Double = (A * ImageData(Idx2, Idx1)) + B
                        If Scaled >= Single.MinValue And Scaled <= Single.MaxValue Then
                            ValueToStore = CType(Scaled, Single)
                        Else
                            ValueToStore = SingleValueInvalid : RetVal += 1
                        End If
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ValueToStore)
                        BytesOut.Write(New Byte() {DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case eBitPix.Double
                Dim ValueToStore As Double = 0
                For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
                    For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                        Dim Scaled As Double = (A * ImageData(Idx2, Idx1)) + B
                        ValueToStore = Scaled
                        Dim DataToWrite As Byte() = BitConverter.GetBytes(ValueToStore)
                        BytesOut.Write(New Byte() {DataToWrite(7), DataToWrite(6), DataToWrite(5), DataToWrite(4), DataToWrite(3), DataToWrite(2), DataToWrite(1), DataToWrite(0)})
                    Next Idx2
                Next Idx1
            Case Else
                Throw New FITSWriterException("Conversion from Double to BitPix <" & CStr(CInt(BitPix)).Trim & "> is not supported!")
        End Select

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

        'Return number of conversion errors
        Return RetVal

    End Function

    '======================================================================================================================================================
    ' Helper functions (private)
    '======================================================================================================================================================

    '''<summary>Get the bytes to store as specified by the FITS standard without any scaling.</summary>
    '''<see cref="http://archive.stsci.edu/fits/fits_standard/node45.html#SECTION001021000000000000000"/>
    Private Shared Function GetBytes_BitPix8(ByVal Value As Byte) As Byte
        Return Value
    End Function

    '''<summary>Get the bytes to store as specified by the FITS standard without any scaling.</summary>
    '''<see cref="http://archive.stsci.edu/fits/fits_standard/node46.html#SECTION001022000000000000000"/>
    Private Shared Function GetBytes_BitPix16(ByVal Value As Int16) As Byte()
        Dim RetVal As Byte() = BitConverter.GetBytes(Value)
        Return New Byte() {RetVal(1), RetVal(0)}
    End Function

    '''<summary>Get the bytes to store as specified by the FITS standard without any scaling.</summary>
    '''<see cref="http://archive.stsci.edu/fits/fits_standard/node47.html"/>
    Private Shared Function GetBytes_BitPix32(ByVal Value As Int32) As Byte()
        Dim RetVal As Byte() = BitConverter.GetBytes(Value)
        Return New Byte() {RetVal(3), RetVal(2), RetVal(1), RetVal(0)}
    End Function

    '''<summary>Get the bytes to store as specified by the FITS standard without any scaling.</summary>
    '''<see cref="http://archive.stsci.edu/fits/fits_standard/node49.html"/>
    Private Shared Function GetBytes_BitPix32f(ByVal Value As Single) As Byte()
        Dim RetVal As Byte() = BitConverter.GetBytes(Value)
        Return New Byte() {RetVal(3), RetVal(2), RetVal(1), RetVal(0)}
    End Function

    '''<summary>Get the bytes to store as specified by the FITS standard without any scaling.</summary>
    '''<see cref="http://archive.stsci.edu/fits/fits_standard/node49.html"/>
    Private Shared Function GetBytes_BitPix64f(ByVal Value As Double) As Byte()
        Dim RetVal As Byte() = BitConverter.GetBytes(Value)
        Return New Byte() {RetVal(7), RetVal(6), RetVal(5), RetVal(4), RetVal(3), RetVal(2), RetVal(1), RetVal(0)}
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------------

    '''<summary>Try to store the passed value in BitPix=8 format.</summary>
    Private Shared Function GetBytes_BitPix8(ByVal Value As Double, ByVal A As Double, ByVal B As Double, ByRef ErrorCount As Integer) As Byte
        Dim Scaled As Double = (A * Value) + B
        If Scaled >= Byte.MinValue And Scaled <= Byte.MaxValue Then
            Return CType(Scaled, Byte)
        Else
            ErrorCount += 1
            Return ByteValueInvalid
        End If
    End Function

    '''<summary>Try to store the passed value in BitPix=16 format.</summary>
    Private Shared Function GetBytes_BitPix16(ByVal Value As Double, ByVal A As Double, ByVal B As Double, ByRef ErrorCount As Integer) As Byte()
        Dim RetVal As Byte() = {}
        Dim Scaled As Double = (A * Value) + B
        If Scaled >= Int16.MinValue And Scaled <= Int16.MaxValue Then
            RetVal = BitConverter.GetBytes(CType(Scaled, Int16))
        Else
            ErrorCount += 1
            RetVal = BitConverter.GetBytes(Int16ValueInvalid)
        End If
        Return New Byte() {RetVal(1), RetVal(0)}
    End Function


    '''<summary>Try to store the passed value in BitPix=32 format.</summary>
    Private Shared Function GetBytes_BitPix32(ByVal Value As Double, ByVal A As Double, ByVal B As Double, ByRef ErrorCount As Integer) As Byte()
        Dim RetVal As Byte() = {}
        Dim Scaled As Double = (A * Value) + B
        If Scaled >= Int32.MinValue And Scaled <= Int32.MaxValue Then
            RetVal = BitConverter.GetBytes(CType(Scaled, Int32))
        Else
            ErrorCount += 1
            RetVal = BitConverter.GetBytes(Int32ValueInvalid)
        End If
        Return New Byte() {RetVal(3), RetVal(2), RetVal(1), RetVal(0)}
    End Function

    '''<summary>Add a custom header to the passed header element.</summary>
    Private Shared Sub AddCustomHeaders(ByRef Header As Collections.Generic.List(Of String), ByRef CustomHeaderElements As Collections.Generic.List(Of String()))
        If IsNothing(CustomHeaderElements) = True Then Exit Sub
        For Each Element As String() In CustomHeaderElements
            If IsNothing(Element) = False Then
                Select Case Element.Length
                    Case 2
                        Header.Add(FormatHeader(Element(0), Element(1)))
                    Case 3
                        Header.Add(FormatHeader(Element(0), Element(1), Element(2)))
                End Select
            End If
        Next Element
    End Sub

    '''<summary>Ensure that the header length is conform with the FITS specification.</summary>
    Private Shared Sub PadHeader(ByRef Header As Collections.Generic.List(Of String))
        Header.Add("END".PadRight(HeaderElementLength))
        If Header.Count Mod HeaderElements <> 0 Then
            Do
                Header.Add(New String(" "c, HeaderElementLength))
            Loop Until Header.Count Mod HeaderElements = 0
        End If
    End Sub

    '''<summary>Format the header according to the FITS standards.</summary>
    Private Shared Function FormatHeader(ByVal Keyword As String, ByVal Value As Integer) As String
        Return FormatHeader(Keyword, Value.ValRegIndep)
    End Function

    '''<summary>Format the header according to the FITS standards.</summary>
    Private Shared Function FormatHeader(ByVal Keyword As String, ByVal Value As String) As String
        Return FormatHeader(Keyword, Value, String.Empty)
    End Function

    '''<summary>Format the header according to the FITS standards.</summary>
    Private Shared Function FormatHeader(ByVal Keyword As String, ByVal Value As String, ByVal Comment As String) As String
        If Keyword.Length > KeywordLength Then Keyword = Keyword.Substring(0, KeywordLength)
        If String.IsNullOrEmpty(Comment) = True Then
            Return (Keyword.PadRight(KeywordLength) & "= " & Value.PadLeft(ValueLength)).PadRight(HeaderElementLength)
        Else
            Return (Keyword.PadRight(KeywordLength) & "= " & Value.PadLeft(ValueLength) & " /" & Comment).PadRight(HeaderElementLength)
        End If
    End Function

    '################################################################################################
    ' TEST ROUTINES
    '################################################################################################

    '''<summary>Write a FITS test file with raw data.</summary>
    '''<remarks>Does work.</remarks>
    Public Shared Sub WriteTestFile_Int8(ByVal FileName As String)

        Dim BitPix As Integer = eBitPix.Byte
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Create test data
        Dim StartValue As Byte = Byte.MinValue
        Dim ImageSize As Integer = 256
        Dim ImageData(ImageSize - 1, ImageSize - 1) As Byte
        Dim Value As Byte = StartValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                ImageData(Idx2, Idx1) = Value
                If Value < Byte.MaxValue Then Value = Value + CType(1, Byte) Else Value = StartValue
            Next Idx2
        Next Idx1

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", "0"))
        Header.Add(FormatHeader("BSCALE", "1"))

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                BytesOut.Write(GetBytes_BitPix8(ImageData(Idx2, Idx1)))
            Next Idx2
        Next Idx1

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '''<summary>Write a FITS test file with raw data.</summary>
    '''<remarks>Does work.</remarks>
    Public Shared Sub WriteTestFile_Int16(ByVal FileName As String)

        Dim BitPix As Integer = eBitPix.Int16
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Create test data
        Dim ImageSize As Integer = 256
        Dim StartValue As Int16 = Int16.MinValue
        Dim ImageData(ImageSize - 1, ImageSize - 1) As Int16
        Dim Value As Int16 = StartValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                ImageData(Idx2, Idx1) = Value
                If Value < Int16.MaxValue Then Value = Value + CType(1, Int16) Else Value = StartValue
            Next Idx2
        Next Idx1

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", "0"))
        Header.Add(FormatHeader("BSCALE", "1"))

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                BytesOut.Write(GetBytes_BitPix16(ImageData(Idx2, Idx1)))
            Next Idx2
        Next Idx1

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '''<summary>Write a FITS test file with raw data.</summary>
    '''<remarks>Does work but not over the full range of 32-bit floating point.</remarks>
    Public Shared Sub WriteTestFile_Int32(ByVal FileName As String)

        Dim BitPix As Integer = eBitPix.Int32
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Create test data
        Dim ImageSize As Integer = 256
        Dim StartValue As Int32 = -1000000
        Dim ImageData(ImageSize - 1, ImageSize - 1) As Int32
        Dim Value As Int32 = StartValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                ImageData(Idx2, Idx1) = Value
                If Value < Int32.MaxValue Then Value = Value + CType(1, Int32) Else Value = StartValue
            Next Idx2
        Next Idx1

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", "0"))
        Header.Add(FormatHeader("BSCALE", "1"))

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                BytesOut.Write(GetBytes_BitPix32(ImageData(Idx2, Idx1)))
            Next Idx2
        Next Idx1

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '''<summary>Write a FITS test file with raw data.</summary>
    '''<remarks>Does work but not over the full range of 32-bit floating point.</remarks>
    Public Shared Sub WriteTestFile_Float32(ByVal FileName As String)

        Dim BitPix As Integer = eBitPix.Single
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Create test data
        Dim ImageSize As Integer = 256
        Dim StartValue As Single = -1000000000.0
        Dim ImageData(ImageSize - 1, ImageSize - 1) As Single
        Dim Value As Single = StartValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                ImageData(Idx2, Idx1) = Value
                If Value < Single.MaxValue Then Value = Value + CType(1000000.0, Single) Else Value = StartValue
            Next Idx2
        Next Idx1

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", "0"))
        Header.Add(FormatHeader("BSCALE", "1"))

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                BytesOut.Write(GetBytes_BitPix32f(ImageData(Idx2, Idx1)))
            Next Idx2
        Next Idx1

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

    '''<summary>Write a FITS test file with raw data.</summary>
    '''<remarks>Does work but not over the full range of 32-bit floating point.</remarks>
    Public Shared Sub WriteTestFile_Float64(ByVal FileName As String)

        Dim BitPix As Integer = eBitPix.Double
        Dim BaseOut As New System.IO.StreamWriter(FileName)
        Dim BytesOut As New System.IO.BinaryWriter(BaseOut.BaseStream)

        'Create test data
        Dim ImageSize As Integer = 256
        Dim StartValue As Double = -1000000000.0
        Dim ImageData(ImageSize - 1, ImageSize - 1) As Double
        Dim Value As Double = StartValue
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                ImageData(Idx2, Idx1) = Value
                If Value < Double.MaxValue Then Value = Value + CType(1000000.0, Double) Else Value = StartValue
            Next Idx2
        Next Idx1

        'Load all header elements
        Dim Header As New Collections.Generic.List(Of String)
        Header.Add(FormatHeader("SIMPLE", "T"))
        Header.Add(FormatHeader("BITPIX", CStr(CInt(BitPix)).Trim))
        Header.Add(FormatHeader("NAXIS", "2"))
        Header.Add(FormatHeader("NAXIS1", ImageData.GetUpperBound(0) + 1))
        Header.Add(FormatHeader("NAXIS2", ImageData.GetUpperBound(1) + 1))
        Header.Add(FormatHeader("BZERO", "0"))
        Header.Add(FormatHeader("BSCALE", "1"))

        'Write header
        PadHeader(Header)
        BaseOut.Write(Join(Header.ToArray, String.Empty))
        BaseOut.Flush()

        'Write content
        For Idx1 As Integer = 0 To ImageData.GetUpperBound(1)
            For Idx2 As Integer = 0 To ImageData.GetUpperBound(0)
                BytesOut.Write(GetBytes_BitPix64f(ImageData(Idx2, Idx1)))
            Next Idx2
        Next Idx1

        'Finish
        BytesOut.Flush()
        BaseOut.Close()

    End Sub

End Class