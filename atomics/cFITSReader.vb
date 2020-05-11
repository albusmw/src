Option Explicit On
Option Strict On

'''<summary>Class to read in FITS data.</summary>
Public Class cFITSReader

    '''<summary>Length of one header element.</summary>
    Const HeaderElementLength As Integer = 80
    '''<summary>Length of a header block - FITS files may contain an integer size of header blocks.</summary>
    Const HeaderBlockSize As Integer = 2880

    '''<summary>Number of header elements per header block.</summary>
    Public Shared ReadOnly HeaderElements As Integer = HeaderBlockSize \ HeaderElementLength

    '''<summary>Path to ipps.dll and ippvm.dll - if not set IPP will not be used.</summary>
    Public Shared Property IPPPath As String = String.Empty

    '''<summary>Instance of Intel IPP library call.</summary>
    Private IntelIPP As cIntelIPP = Nothing

    Private Interface IByteConverter
        Function Convert(ByRef Bytes() As Byte, ByVal Offset As Integer) As Double
    End Interface

    Public Class cByteConverter_Byte : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Return RawData(Offset)
        End Function
    End Class

    Public Class cByteConverter_Int16 : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Return BitConverter.ToInt16({RawData(Offset + 1), RawData(Offset)}, 0)
        End Function
    End Class

    Public Class cByteConverter_Int32 : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Return BitConverter.ToInt32({RawData(Offset + 3), RawData(Offset + 2), RawData(Offset + 1), RawData(Offset)}, 0)
        End Function
    End Class

    Public Class cByteConverter_Int32_Fast : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Dim Val1 As Int32 = (CInt(RawData(0)) << 24) + (CInt(RawData(Offset + 1)) << 16) + (CInt(RawData(Offset + 2)) << 8) + (CInt(RawData(Offset + 3)))
            'Dim Val2 As Int32 = BitConverter.ToInt32({RawData(Offset + 3), RawData(Offset + 2), RawData(Offset + 1), RawData(Offset)}, 0)
            Return Val1
        End Function
    End Class

    Public Class cByteConverter_Single : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Return BitConverter.ToSingle({RawData(Offset + 3), RawData(Offset + 2), RawData(Offset + 1), RawData(Offset)}, 0)
        End Function
    End Class

    Public Class cByteConverter_Double : Implements IByteConverter
        Public Function Convert(ByRef RawData() As Byte, ByVal Offset As Integer) As Double Implements IByteConverter.Convert
            Return BitConverter.ToDouble({RawData(Offset + 7), RawData(Offset + 6), RawData(Offset + 5), RawData(Offset + 4), RawData(Offset + 3), RawData(Offset + 2), RawData(Offset + 1), RawData(Offset)}, 0)
        End Function
    End Class

    Private FITSHeaderParser As cFITSHeaderParser = Nothing

    Public Sub New()
        IntelIPP = New cIntelIPP(System.IO.Path.Combine(IPPPath, "ipps.dll"), System.IO.Path.Combine(IPPPath, "ippvm.dll"), System.IO.Path.Combine(IPPPath, "ippi.dll"))
    End Sub

    Public Sub ReadIn(ByVal FileName As String, ByRef ImageData(,) As Double)
        ReadIn(FileName, True, ImageData, New System.Drawing.Point() {})
    End Sub

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseBZeroScale">Use the BZERO and BSCALE value within the file for scaling - if OFF omit the data.</param>
    '''<param name="ImageData">Loaded image data processed by BZERO and BSCALE - if PointsToRead is passed, the matrix is 1xN where N is the number of entries in PointsToRead.</param>
    Public Sub ReadIn(ByVal FileName As String, ByVal UseBZeroScale As Boolean, ByRef ImageData(,) As Double)
        ReadIn(FileName, UseBZeroScale, ImageData, New System.Drawing.Point() {})
    End Sub

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseBZeroScale">Use the BZERO and BSCALE value within the file for scaling - if OFF omit the data.</param>
    '''<param name="ImageData">Loaded image data processed by BZERO and BSCALE - if PointsToRead is passed, the matrix is 1xN where N is the number of entries in PointsToRead.</param>
    '''<param name="PointsToRead">Vector of points to read on - pass an empty vector to read all values and generate a matrix for ImageData.</param>
    Public Sub ReadIn(ByVal FileName As String, ByVal UseBZeroScale As Boolean, ByRef ImageData(,) As Double, ByVal PointsToRead As System.Drawing.Point())

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        ReadDataContent(FileName, DataStartPos, ImageData, FITSHeaderParser.BitPix, UseBZeroScale, FITSHeaderParser.Width, FITSHeaderParser.Height, PointsToRead)

    End Sub

    '================================================================================================================================================================
    ' Read UInt8 data (data are read to an UInt16 matrix)
    '================================================================================================================================================================

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    '''<remarks>Tested and works.</remarks>
    Public Function ReadInUInt8(ByVal FileName As String, ByVal UseIPP As Boolean) As UInt16(,)

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        Return ReadDataContentUInt8(FileName, DataStartPos, UseIPP)

    End Function

    '''<summary>Read FITS data from the passed file - only in case BitPix is 16.</summary>
    Private Function ReadDataContentUInt8(ByVal FileName As String, ByVal StartPosition As Integer, ByVal UseIPP As Boolean) As UInt16(,)

        Dim BytePerPixel As Integer = FITSHeaderParser.BytesPerSample

        'Delete content and exit if format is wrong
        If FITSHeaderParser.BitPix <> 8 Then Return New UInt16(,) {}

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = StartPosition

        'Read complete block
        Dim ImageData(FITSHeaderParser.Width - 1, FITSHeaderParser.Height - 1) As UInt16
        Dim Bytes((FITSHeaderParser.Width * FITSHeaderParser.Height * BytePerPixel) - 1) As Byte
        Bytes = DataReader.ReadBytes(Bytes.Length)
        Dim BytesPtr As Integer = 0
        For H As Integer = 0 To FITSHeaderParser.Height - 1
            For W As Integer = 0 To FITSHeaderParser.Width - 1
                ImageData(W, H) = Bytes(BytesPtr)
                BytesPtr += BytePerPixel
            Next W
        Next H

        'Close data stream
        DataReader.Close()

        Return ImageData

    End Function

    '================================================================================================================================================================
    ' Read UInt16 data
    '================================================================================================================================================================

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    '''<remarks>Tested and works.</remarks>
    Public Function ReadInUInt16(ByVal FileName As String, ByVal UseIPP As Boolean) As UInt16(,)
        Return ReadInUInt16(FileName, UseIPP, -1, -1, -1, -1)
    End Function

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    '''<param name="XOffset">0-based X start offset - use -1 to ignore.</param>
    '''<param name="XWidth">Width [pixel] to read in.</param>
    '''<param name="YOffset">0-based Y start offset - use -1 to ignore.</param>
    '''<param name="YHeight">Height [pixel] to read in.</param>
    '''<remarks>Tested and works.</remarks>
    Public Function ReadInUInt16(ByVal FileName As String, ByVal UseIPP As Boolean, ByVal XOffset As Integer, ByVal XWidth As Integer, ByVal YOffset As Integer, ByVal YHeight As Integer) As UInt16(,)
        Return ReadInUInt16(FileName, -1, UseIPP, XOffset, XWidth, YOffset, YHeight)
    End Function

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="DataStartPosToUse">OVerrided data start index (used e.g. to process NAXIS3>1 pictures).</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    '''<param name="XOffset">0-based X start offset - use -1 to ignore.</param>
    '''<param name="XWidth">Width [pixel] to read in.</param>
    '''<param name="YOffset">0-based Y start offset - use -1 to ignore.</param>
    '''<param name="YHeight">Height [pixel] to read in.</param>
    '''<remarks>Tested and works.</remarks>
    Public Function ReadInUInt16(ByVal FileName As String, ByVal DataStartPosToUse As Integer, ByVal UseIPP As Boolean) As UInt16(,)
        Return ReadInUInt16(FileName, DataStartPosToUse, UseIPP, -1, -1, -1, -1)
    End Function

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="DataStartPosToUse">OVerrided data start index (used e.g. to process NAXIS3>1 pictures).</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    '''<param name="XOffset">0-based X start offset - use -1 to ignore.</param>
    '''<param name="XWidth">Width [pixel] to read in.</param>
    '''<param name="YOffset">0-based Y start offset - use -1 to ignore.</param>
    '''<param name="YHeight">Height [pixel] to read in.</param>
    '''<remarks>Tested and works.</remarks>
    Public Function ReadInUInt16(ByVal FileName As String, ByVal DataStartPosToUse As Integer, ByVal UseIPP As Boolean, ByVal XOffset As Integer, ByVal XWidth As Integer, ByVal YOffset As Integer, ByVal YHeight As Integer) As UInt16(,)

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        If DataStartPosToUse > -1 Then DataStartPos = DataStartPosToUse
        Return ReadDataContentUInt16(FileName, DataStartPos, UseIPP, XOffset, XWidth, YOffset, YHeight)

    End Function

    '''<summary>Read FITS data from the passed file - only in case BitPix is 16.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="DataStartPosition">Position of the data start.</param>
    '''<param name="XOffset">0-based X start offset - use -1 to ignore.</param>
    '''<param name="XWidth">Width [pixel] to read in.</param>
    '''<param name="YOffset">0-based Y start offset - use -1 to ignore.</param>
    '''<param name="YHeight">Height [pixel] to read in.</param>
    Public Function ReadDataContentUInt16(ByVal FileName As String, ByVal DataStartPosition As Integer, ByVal UseIPP As Boolean, ByVal XOffset As Integer, ByVal XWidth As Integer, ByVal YOffset As Integer, ByVal YHeight As Integer) As UInt16(,)

        Dim BytePerPixel As Integer = 2

        'Exit if format is wrong
        If FITSHeaderParser.BitPix <> 16 Then Return New UInt16(,) {}

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPosition

        Dim ImageData(,) As UInt16 = {}
        Dim Bytes() As Byte = {}
        If XWidth = -1 And YHeight = -1 Then
            'Read complete block
            ReDim ImageData(FITSHeaderParser.Width - 1, FITSHeaderParser.Height - 1)
            Bytes = DataReader.ReadBytes((FITSHeaderParser.Width * FITSHeaderParser.Height * BytePerPixel) - 1)
        Else
            'Read only a part
            ReDim ImageData(XWidth - 1, YHeight - 1)
            ReDim Bytes(CInt((ImageData.LongLength * BytePerPixel) - 1))
            Dim BytesPtr As Integer = 0
            For H As Integer = 0 To ImageData.GetUpperBound(1)
                Dim PixelOffset As Integer = (((YOffset + H) * FITSHeaderParser.Width) + XOffset)
                DataReader.BaseStream.Position = DataStartPosition + (BytePerPixel * PixelOffset)
                    Dim Part() As Byte = DataReader.ReadBytes((ImageData.GetUpperBound(0) + 1) * BytePerPixel)
                    Part.CopyTo(Bytes, BytesPtr)
                BytesPtr += Part.Length
            Next H
        End If


        If UseIPP = False Or FITSHeaderParser.BZERO <> 32768 Or FITSHeaderParser.BSCALE <> 1 Then
            'VB implementation
            Dim BytesPtr As Integer = 0
            For H As Integer = 0 To ImageData.GetUpperBound(1)
                For W As Integer = 0 To ImageData.GetUpperBound(0)
                    ImageData(W, H) = CUShort(BitConverter.ToInt16({Bytes(BytesPtr + 1), Bytes(BytesPtr)}, 0) + FITSHeaderParser.BZERO)
                    BytesPtr += BytePerPixel
                Next W
            Next H
        Else
            'IPP implementation
            Dim IPPStatus As New List(Of cIntelIPP.IppStatus)
            IPPStatus.Add(IntelIPP.Transpose(Bytes, ImageData))
            IPPStatus.Add(IntelIPP.SwapBytes(ImageData))
            IPPStatus.Add(IntelIPP.XorC(ImageData, &H8000))
        End If

        'Close data stream
        DataReader.Close()

        Return ImageData

    End Function

    '================================================================================================================================================================
    ' Read Int32 data
    '================================================================================================================================================================

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="ImageData">Loaded image data processed by BZERO and BSCALE - if PointsToRead is passed, the matrix is 1xN where N is the number of entries in PointsToRead.</param>
    Public Sub ReadIn(ByVal FileName As String, ByRef ImageData(,) As Int32)

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        ReadDataContent(FileName, DataStartPos, ImageData, FITSHeaderParser.BitPix, FITSHeaderParser.Width, FITSHeaderParser.Height)

    End Sub

    '''<summary>Read FITS data from the passed file - only in case BitPix is 32.</summary>
    Private Sub ReadDataContent(ByVal FileName As String, ByVal DataStartPos As Integer, ByRef ImageData(,) As Int32, ByVal BitPix As Integer, ByVal Width As Integer, ByVal Height As Integer)

        Dim Stopper As New Stopwatch
        Stopper.Reset() : Stopper.Start()

        'Delete content and exit if format is wrong
        ImageData = {}
        If BitPix <> 32 Then Exit Sub

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPos

        'Read all data
        ReDim ImageData(Width - 1, Height - 1)
        For H As Integer = 0 To Height - 1
            For W As Integer = 0 To Width - 1
                ImageData(W, H) = DataReader.ReadInt32
            Next W
        Next H

        'Convert format
        IntelIPP.SwapBytes(ImageData)

        'Close data stream
        DataReader.Close()

        Stopper.Stop()
        Debug.Print("Reading FITS data content took " & Stopper.ElapsedMilliseconds.ValRegIndep & " ms")

    End Sub

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    Public Function ReadInInt32(ByVal FileName As String, ByVal UseIPP As Boolean) As Int32(,)

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        Return ReadDataContentInt32(FileName, DataStartPos, UseIPP)

    End Function

    '''<summary>Read FITS data from the passed file - only in case BitPix is 32.</summary>
    Private Function ReadDataContentInt32(ByVal FileName As String, ByVal DataStartPos As Integer, ByVal UseIPP As Boolean) As Int32(,)

        Dim BytePerPixel As Integer = 4

        'Delete content and exit if format is wrong
        If FITSHeaderParser.BitPix <> 32 Then Return New Int32(,) {}

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPos

        'Read complete block
        Dim ImageData(FITSHeaderParser.Width - 1, FITSHeaderParser.Height - 1) As Int32
        Dim Bytes((FITSHeaderParser.Width * FITSHeaderParser.Height * BytePerPixel) - 1) As Byte
        Bytes = DataReader.ReadBytes(Bytes.Length)
        If UseIPP = False Then
            'VB implementation
            Dim FileValue As Int32 = 0
            Dim BytesPtr As Integer = 0
            For H As Integer = 0 To FITSHeaderParser.Height - 1
                For W As Integer = 0 To FITSHeaderParser.Width - 1
                    FileValue = BitConverter.ToInt32({Bytes(BytesPtr + 3), Bytes(BytesPtr + 2), Bytes(BytesPtr + 1), Bytes(BytesPtr)}, 0)
                    ImageData(W, H) = CInt((FITSHeaderParser.BSCALE * FileValue) + FITSHeaderParser.BZERO)
                    BytesPtr += BytePerPixel
                Next W
            Next H
        Else
            Dim IPPStatus As New List(Of cIntelIPP.IppStatus)
            IPPStatus.Add(IntelIPP.Transpose(Bytes, ImageData))
            IPPStatus.Add(IntelIPP.SwapBytes(ImageData))
        End If

        'Close data stream
        DataReader.Close()

        Return ImageData

    End Function

    '================================================================================================================================================================
    ' Read Float32 data
    '================================================================================================================================================================

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="UseIPP">Use the Intel IPP (if found) for processing.</param>
    Public Function ReadInFloat32(ByVal FileName As String, ByVal UseIPP As Boolean) As Single(,)

        'Read in header and get data start position
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        FITSHeaderParser = New cFITSHeaderParser(ReadHeader(BaseIn, DataStartPos))
        BaseIn.Close()

        'Read data content
        Return ReadDataContentFloat32(FileName, DataStartPos, UseIPP)

    End Function

    '''<summary>Read FITS data from the passed file - only in case BitPix is 32.</summary>
    Private Function ReadDataContentFloat32(ByVal FileName As String, ByVal DataStartPos As Integer, ByVal UseIPP As Boolean) As Single(,)

        Dim BytePerPixel As Integer = 4

        'Delete content and exit if format is wrong
        If FITSHeaderParser.BitPix <> -32 Then Return New Single(,) {}

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPos

        'Read complete block
        Dim ImageData(FITSHeaderParser.Width - 1, FITSHeaderParser.Height - 1) As Single
        Dim Bytes((FITSHeaderParser.Width * FITSHeaderParser.Height * BytePerPixel) - 1) As Byte
        Bytes = DataReader.ReadBytes(Bytes.Length)
        'If UseIPP = False Then
        'VB implementation
        Dim FileValue As Single = 0
        Dim BytesPtr As Integer = 0
        If FITSHeaderParser.BSCALE = 1 And FITSHeaderParser.BZERO = 0 Then
            'Direct copy as no scaling is applied
            For H As Integer = 0 To FITSHeaderParser.Height - 1
                For W As Integer = 0 To FITSHeaderParser.Width - 1
                    ImageData(W, H) = BitConverter.ToSingle({Bytes(BytesPtr + 3), Bytes(BytesPtr + 2), Bytes(BytesPtr + 1), Bytes(BytesPtr)}, 0)
                    BytesPtr += BytePerPixel
                Next W
            Next H
        Else
            'BSCALE and BZERO must be applied
            For H As Integer = 0 To FITSHeaderParser.Height - 1
                For W As Integer = 0 To FITSHeaderParser.Width - 1
                    FileValue = BitConverter.ToSingle({Bytes(BytesPtr + 3), Bytes(BytesPtr + 2), Bytes(BytesPtr + 1), Bytes(BytesPtr)}, 0)
                    ImageData(W, H) = CSng((FITSHeaderParser.BSCALE * FileValue) + FITSHeaderParser.BZERO)
                    BytesPtr += BytePerPixel
                Next W
            Next H
        End If
        'Else
        'Dim IPPStatus As New List(Of cIntelIPP.IppStatus)
        'IPPStatus.Add(IntelIPP.Transpose(Bytes, ImageData))
        'IPPStatus.Add(IntelIPP.SwapBytes(ImageData))
        'End If

        'Close data stream
        DataReader.Close()

        Return ImageData

    End Function

    '================================================================================================================================================================
    ' Read any data to an Double matrix
    '================================================================================================================================================================

    Private Sub ReadDataContent(ByVal FileName As String, ByVal DataStartPos As Integer, ByRef ImageData(,) As Double, ByVal BitPix As Integer, ByVal UseBZeroScale As Boolean, ByVal Width As Integer, ByVal Height As Integer, ByVal PointsToRead As System.Drawing.Point())

        'Performance comments:
        ' - Apply BZERO and BSCALE just then reading in - array manipulation access seems to be slow
        ' - Reading in all data at once is faster but needs more memory
        'Improvements possible:
        ' - If only some points need to be read in, direct stream access will be faster (and less memory consuming) compared to "read all"

        Dim Stopper As New Stopwatch
        Stopper.Reset() : Stopper.Start()

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPos

        'Get number of bytes per value and the converter to be used
        Dim Converter As IByteConverter = Nothing
        Dim BytesPerSample As Integer = -1
        Select Case BitPix
            Case 8
                BytesPerSample = 1 : Converter = New cByteConverter_Byte
            Case 16
                BytesPerSample = 2 : Converter = New cByteConverter_Int16
            Case 32
                BytesPerSample = 4 : Converter = New cByteConverter_Int32_Fast
            Case -32
                BytesPerSample = 4 : Converter = New cByteConverter_Single
            Case -64
                BytesPerSample = 8 : Converter = New cByteConverter_Double
        End Select

        'Set image and buffer data
        Dim AllRawData As Byte() : ReDim AllRawData((Height * Width * BytesPerSample) - 1)
        DataReader.Read(AllRawData, 0, AllRawData.Length)
        Dim RawDataPtr As Integer = 0

        'Check if all data should be read
        Dim ReadAllData As Boolean = False
        If IsNothing(PointsToRead) = True Then ReadAllData = True Else If PointsToRead.Length = 0 Then ReadAllData = True

        'Read data - first select if all points must be read or only a dedicated number of points
        If ReadAllData Then
            'Read all data
            ReDim ImageData(Width - 1, Height - 1)
            If (FITSHeaderParser.BZERO <> 0 Or FITSHeaderParser.BSCALE <> 1) And UseBZeroScale = True Then
                'Scaling
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = FITSHeaderParser.BZERO + (FITSHeaderParser.BSCALE * Converter.Convert(AllRawData, RawDataPtr))
                        RawDataPtr += BytesPerSample
                    Next W
                Next H
            Else
                'No scaling
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = Converter.Convert(AllRawData, RawDataPtr)
                        RawDataPtr += BytesPerSample
                    Next W
                Next H
            End If

        Else
            'Read only specific points
            ReDim ImageData(PointsToRead.Length - 1, 0)
            If (FITSHeaderParser.BZERO <> 0 Or FITSHeaderParser.BSCALE <> 1) And UseBZeroScale = True Then
                'Scaling
                For Idx As Integer = 0 To PointsToRead.Length - 1
                    RawDataPtr = BytesPerSample * ((PointsToRead(Idx).Y * Width) + PointsToRead(Idx).X)
                    ImageData(Idx, 0) = FITSHeaderParser.BZERO + (FITSHeaderParser.BSCALE * Converter.Convert(AllRawData, RawDataPtr))
                Next Idx
            Else
                'No scaling
                For Idx As Integer = 0 To PointsToRead.Length - 1
                    RawDataPtr = BytesPerSample * ((PointsToRead(Idx).Y * Width) + PointsToRead(Idx).X)
                    ImageData(Idx, 0) = Converter.Convert(AllRawData, RawDataPtr)
                Next Idx
            End If
        End If

        'Close data stream
        DataReader.Close()

        Stopper.Stop()
        Debug.Print("Reading FITS data content took " & Stopper.ElapsedMilliseconds.ValRegIndep & " ms")

    End Sub

    '''<summary>Read FITS data from the passed file, and do not apply any scaling indicated.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="ImageData">Loaded image data as-is.</param>
    Public Sub ReadInRaw(ByVal FileName As String, ByRef ImageData(,) As Int32)

        'TODO: Read-in start offset seems to be incorrect
        Dim BaseIn As New System.IO.StreamReader(FileName)
        Dim DataStartPos As Integer = -1
        Dim HeaderEntries As List(Of cFITSHeaderParser.sHeaderElement) = ReadHeader(BaseIn, DataStartPos)
        BaseIn.Close()

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartPos

        'Set image and buffer add data
        Dim PtrStepping As Integer = FITSHeaderParser.BitPix \ 8
        Dim AllRawData As Byte() : ReDim AllRawData((FITSHeaderParser.Height * FITSHeaderParser.Width * PtrStepping) - 1)
        DataReader.Read(AllRawData, 0, AllRawData.Length)

        'Read all data
        ReDim ImageData(FITSHeaderParser.Width - 1, FITSHeaderParser.Height - 1)
        Dim RawDataPtr As Integer = 0
        Select Case FITSHeaderParser.BitPix
            Case 8
                For H As Integer = 0 To FITSHeaderParser.Height - 1
                    For W As Integer = 0 To FITSHeaderParser.Width - 1
                        ImageData(W, H) = AllRawData(RawDataPtr)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
            Case 16
                For H As Integer = 0 To FITSHeaderParser.Height - 1
                    For W As Integer = 0 To FITSHeaderParser.Width - 1
                        ImageData(W, H) = BitConverter.ToInt16({AllRawData(RawDataPtr + 1), AllRawData(RawDataPtr)}, 0)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
            Case 32
                For H As Integer = 0 To FITSHeaderParser.Height - 1
                    For W As Integer = 0 To FITSHeaderParser.Width - 1
                        ImageData(W, H) = BitConverter.ToInt32({AllRawData(RawDataPtr + 3), AllRawData(RawDataPtr + 2), AllRawData(RawDataPtr + 1), AllRawData(RawDataPtr)}, 0)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
        End Select

    End Sub

    '''<summary>Change pixel in the passed FIT file.</summary>
    '''<param name="FileName">FITS file to modify.</param>
    '''<param name="PointToWrite">List of points to be modified.</param>
    '''<param name="FixValues">Values to use for modification.</param>
    Public Sub FixSample(ByVal FileName As String, ByVal DataStartPos As Integer, ByRef PointToWrite As List(Of System.Drawing.Point), ByVal FixValues As Int16())
        Dim DataWriter As New System.IO.BinaryWriter(System.IO.File.OpenWrite(FileName))
        For Idx As Integer = 0 To PointToWrite.Count - 1
            Dim BytesToWrite As Byte() = BitConverter.GetBytes(FixValues(Idx))
            Dim PixelOffsetPtr As Integer = FITSHeaderParser.BytesPerSample * ((PointToWrite(Idx).Y * FITSHeaderParser.Width) + PointToWrite(Idx).X)
            DataWriter.Seek(DataStartPos + PixelOffsetPtr, IO.SeekOrigin.Begin)
            DataWriter.Write(BytesToWrite(1)) : DataWriter.Write(BytesToWrite(0))
        Next Idx
        DataWriter.Flush()
        DataWriter.Close()
    End Sub

    '==================================================================================================================
    'INTERNAL HELPER FUNCTIONS
    '==================================================================================================================

    '''<summary>Get a list of all header elements.</summary>
    '''<param name="BaseIn">Stream to read data in from.</param>
    '''<param name="DataStartPos">0-based index where the data start.</param>
    '''<returns>List of header elements.</returns>
    Private Function ReadHeader(ByRef BaseIn As System.IO.StreamReader, ByRef DataStartPos As Integer) As List(Of cFITSHeaderParser.sHeaderElement)

        Dim EndReached As Boolean = False
        Dim RetVal As New List(Of cFITSHeaderParser.sHeaderElement)
        Dim BlocksRead As Integer = 0

        Do
            Dim Buffer As Char() : ReDim Buffer(HeaderBlockSize - 1)
            Dim Header As String = String.Empty
            Do
                'If the header is empty but the END tag is not yet found, read again
                If Header.Length = 0 Then
                    BaseIn.ReadBlock(Buffer, 0, Buffer.Length)
                    BlocksRead += 1
                    Header = New String(Buffer)
                End If
                Dim NewLine As String = Header.Substring(0, HeaderElementLength)
                Header = Header.Substring(HeaderElementLength)
                If NewLine.StartsWith("END") Then
                    EndReached = True
                Else
                    Dim NewHeaderElement As cFITSHeaderParser.sHeaderElement = Nothing
                    NewHeaderElement.Keyword = cFITSHeaderParser.GetKeywordEnum(NewLine.Substring(0, 8).Trim)
                    NewLine = NewLine.Substring(10)
                    Dim Splitted As String() = Split(NewLine, "/")
                    NewHeaderElement.Value = Splitted(0).Trim
                    If Splitted.Length > 1 Then NewHeaderElement.Comment = Splitted(1)
                    RetVal.Add(NewHeaderElement)
                End If
            Loop Until EndReached = True
        Loop Until EndReached = True

        DataStartPos = BlocksRead * HeaderBlockSize
        Return RetVal

    End Function

    '''<summary>Convert the passed Int32 value to a UInt32 value for FITS files with BScale=1 and BOffset=32768.</summary>
    Private Function FITSToUnsigned(ByRef RawData As Int32) As UInt32
        Dim Bytes As Byte() = BitConverter.GetBytes(RawData)
        Dim Int32Value As Int32 = BitConverter.ToInt32({Bytes(3), Bytes(2), Bytes(1), Bytes(0)}, 0)
        Return CUInt(Int32Value + 32768)
    End Function

End Class
