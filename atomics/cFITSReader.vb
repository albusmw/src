Option Explicit On
Option Strict On

Public Class cFITSReader

    '''<summary>Length of one header element.</summary>
    Const HeaderElementLength As Integer = 80
    '''<summary>Length of a header block - FITS files may contain an integer size of header blocks.</summary>
    Const HeaderBlockSize As Integer = 2880
    '''<summary>Number of header elements per header block.</summary>
    Public Shared ReadOnly HeaderElements As Integer = HeaderBlockSize \ HeaderElementLength

    '''<summary>Path to ipps.dll and ippvm.dll - if not set IPP will not be used.</summary>
    Public Property IPPPath As String = String.Empty

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

    Public Structure sHeaderElement
        Public Element As String
        Public Value As String
        Public Comment As String
    End Structure

    Private Class cMyProps
        Public BitPix As Integer = 0
        Public BZERO As Double = 0.0
        Public BSCALE As Double = 1.0
        Public Width As Integer = -1
        Public Height As Integer = -1
        Public ColorValues As Integer = 0
        Public BytesPerSample As Integer = -1
        Public DataStartIdx As Integer = -1
    End Class
    Private MyProps As New cMyProps

    Public ReadOnly Property BitPix() As Integer
        Get
            Return MyProps.BitPix
        End Get
    End Property

    Public ReadOnly Property BZERO() As Double
        Get
            Return MyProps.BZERO
        End Get
    End Property


    Public ReadOnly Property BSCALE() As Double
        Get
            Return MyProps.BSCALE
        End Get
    End Property


    Public ReadOnly Property Width() As Integer
        Get
            Return MyProps.Width
        End Get
    End Property


    Public ReadOnly Property Height() As Integer
        Get
            Return MyProps.Height
        End Get
    End Property


    Public ReadOnly Property BytesPerSample() As Integer
        Get
            Return MyProps.BytesPerSample
        End Get
    End Property


    ''' <summary>0-based index of the first image data within the file.</summary>
    Public ReadOnly Property DataStartIdx() As Integer
        Get
            Return MyProps.DataStartIdx
        End Get
    End Property

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

        'TODO: Read-in start offset seems to be incorrect
        Dim BaseIn As New System.IO.StreamReader(FileName)

        'Read header elements
        Dim HeaderEntries As List(Of sHeaderElement) = ReadHeader(BaseIn)

        'Calculate data stream properties
        Dim StartOffset As Long = BaseIn.BaseStream.Position
        Dim StreamLength As Long = BaseIn.BaseStream.Length
        Dim TotalByte As Long = StreamLength - StartOffset
        BaseIn.Close()

        'Read data content
        ReadDataContent(FileName, MyProps.DataStartIdx, ImageData, BitPix, UseBZeroScale, Width, Height, PointsToRead)

    End Sub

    '''<summary>Read FITS data from the passed file.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="ImageData">Loaded image data processed by BZERO and BSCALE - if PointsToRead is passed, the matrix is 1xN where N is the number of entries in PointsToRead.</param>
    Public Sub ReadIn(ByVal FileName As String, ByRef ImageData(,) As Int32)

        'TODO: Read-in start offset seems to be incorrect
        Dim BaseIn As New System.IO.StreamReader(FileName)

        'Read header elements
        Dim HeaderEntries As List(Of sHeaderElement) = ReadHeader(BaseIn)

        'Calculate data stream properties
        Dim StartOffset As Long = BaseIn.BaseStream.Position
        Dim StreamLength As Long = BaseIn.BaseStream.Length
        Dim TotalByte As Long = StreamLength - StartOffset
        BaseIn.Close()

        'Read data content
        ReadDataContent(FileName, MyProps.DataStartIdx, ImageData, BitPix, Width, Height)

    End Sub

    '''<summary>Read FITS data from the passed file - only in case BitPix is 32.</summary>
    Private Sub ReadDataContent(ByVal FileName As String, ByVal StartPosition As Integer, ByRef ImageData(,) As Int32, ByVal BitPix As Integer, ByVal Width As Integer, ByVal Height As Integer)

        Dim Stopper As New Stopwatch
        Stopper.Reset() : Stopper.Start()

        'Delete content and exit if format is wrong
        ImageData = {}
        If BitPix <> 32 Then Exit Sub

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartIdx

        'Read all data
        ReDim ImageData(Width - 1, Height - 1)
        For H As Integer = 0 To Height - 1
            For W As Integer = 0 To Width - 1
                ImageData(W, H) = DataReader.ReadInt32
            Next W
        Next H

        'Convert format
        Dim X As New cIntelIPP(IPPPath & "ipps.dll", IPPPath & "ippvm.dll")
        X.SwapBytes(ImageData)

        'Close data stream
        DataReader.Close()

        Stopper.Stop()
        Debug.Print("Reading FITS data content took " & Stopper.ElapsedMilliseconds.ToString.Trim & " ms")

    End Sub

    Private Sub ReadDataContent(ByVal FileName As String, ByVal StartPosition As Integer, ByRef ImageData(,) As Double, ByVal BitPix As Integer, ByVal UseBZeroScale As Boolean, ByVal Width As Integer, ByVal Height As Integer, ByVal PointsToRead As System.Drawing.Point())

        'Performance comments:
        ' - Apply BZERO and BSCALE just then reading in - array manipulation access seems to be slow
        ' - Reading in all data at once is faster but needs more memory
        'Improvements possible:
        ' - If only some points need to be read in, direct stream access will be faster (and less memory consuming) compared to "read all"

        Dim Stopper As New Stopwatch
        Stopper.Reset() : Stopper.Start()

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartIdx

        'Get number of bytes per value and the converter to be used
        Dim Converter As IByteConverter = Nothing
        Select Case BitPix
            Case 8
                MyProps.BytesPerSample = 1 : Converter = New cByteConverter_Byte
            Case 16
                MyProps.BytesPerSample = 2 : Converter = New cByteConverter_Int16
            Case 32
                MyProps.BytesPerSample = 4 : Converter = New cByteConverter_Int32_Fast
            Case -32
                MyProps.BytesPerSample = 4 : Converter = New cByteConverter_Single
            Case -64
                MyProps.BytesPerSample = 8 : Converter = New cByteConverter_Double
        End Select

        'Set image and buffer data
        Dim AllRawData As Byte() : ReDim AllRawData((Height * Width * MyProps.BytesPerSample) - 1)
        DataReader.Read(AllRawData, 0, AllRawData.Length)
        Dim RawDataPtr As Integer = 0

        'Check if all data should be read
        Dim ReadAllData As Boolean = False
        If IsNothing(PointsToRead) = True Then ReadAllData = True Else If PointsToRead.Length = 0 Then ReadAllData = True

        'Read data - first select if all points must be read or only a dedicated number of points
        If ReadAllData Then
            'Read all data
            ReDim ImageData(Width - 1, Height - 1)
            If (BZERO <> 0 Or BSCALE <> 1) And UseBZeroScale = True Then
                'Scaling
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = BZERO + (BSCALE * Converter.Convert(AllRawData, RawDataPtr))
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
            If (BZERO <> 0 Or BSCALE <> 1) And UseBZeroScale = True Then
                'Scaling
                For Idx As Integer = 0 To PointsToRead.Length - 1
                    RawDataPtr = BytesPerSample * ((PointsToRead(Idx).Y * Width) + PointsToRead(Idx).X)
                    ImageData(Idx, 0) = BZERO + (BSCALE * Converter.Convert(AllRawData, RawDataPtr))
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
        Debug.Print("Reading FITS data content took " & Stopper.ElapsedMilliseconds.ToString.Trim & " ms")

    End Sub

    '''<summary>Read FITS data from the passed file, and do not apply any scaling indicated.</summary>
    '''<param name="FileName">File name to load FITS data from.</param>
    '''<param name="ImageData">Loaded image data as-is.</param>
    Public Sub ReadInRaw(ByVal FileName As String, ByRef ImageData(,) As Int32)

        'TODO: Read-in start offset seems to be incorrect
        Dim BaseIn As New System.IO.StreamReader(FileName)

        'Read header elements
        Dim HeaderEntries As List(Of sHeaderElement) = ReadHeader(BaseIn)

        'Calculate data stream properties
        Dim StartOffset As Long = BaseIn.BaseStream.Position
        Dim StreamLength As Long = BaseIn.BaseStream.Length
        Dim TotalByte As Long = StreamLength - StartOffset
        BaseIn.Close()

        'Open reader and position to start
        Dim DataReader As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))
        DataReader.BaseStream.Position = DataStartIdx

        'Set image and buffer add data
        Dim PtrStepping As Integer = BitPix \ 8
        Dim AllRawData As Byte() : ReDim AllRawData((Height * Width * PtrStepping) - 1)
        DataReader.Read(AllRawData, 0, AllRawData.Length)

        'Read all data
        ReDim ImageData(Width - 1, Height - 1)
        Dim RawDataPtr As Integer = 0
        Select Case BitPix
            Case 8
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = AllRawData(RawDataPtr)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
            Case 16
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = BitConverter.ToInt16({AllRawData(RawDataPtr + 1), AllRawData(RawDataPtr)}, 0)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
            Case 32
                For H As Integer = 0 To Height - 1
                    For W As Integer = 0 To Width - 1
                        ImageData(W, H) = BitConverter.ToInt32({AllRawData(RawDataPtr + 3), AllRawData(RawDataPtr + 2), AllRawData(RawDataPtr + 1), AllRawData(RawDataPtr)}, 0)
                        RawDataPtr += PtrStepping
                    Next W
                Next H
        End Select

    End Sub

    '''<summary>Change entries in the passed FileName FIT file.</summary>
    '''<param name="FileName">File to modify.</param>
    '''<param name="PointToWrite">List of points to be modified.</param>
    '''<param name="FixValues">Values to use for modification.</param>
    Public Sub FixSample(ByVal FileName As String, ByRef PointToWrite As List(Of System.Drawing.Point), ByVal FixValues As Int16())
        Dim DataWriter As New System.IO.BinaryWriter(System.IO.File.OpenWrite(FileName))
        For Idx As Integer = 0 To PointToWrite.Count - 1
            Dim BytesToWrite As Byte() = BitConverter.GetBytes(FixValues(Idx))
            Dim PixelOffsetPtr As Integer = BytesPerSample * ((PointToWrite(Idx).Y * Width) + PointToWrite(Idx).X)
            DataWriter.Seek(DataStartIdx + PixelOffsetPtr, IO.SeekOrigin.Begin)
            DataWriter.Write(BytesToWrite(1)) : DataWriter.Write(BytesToWrite(0))
        Next Idx
        DataWriter.Flush()
        DataWriter.Close()
    End Sub

    '==================================================================================================================
    'INTERNAL HELPER FUNCTIONS
    '==================================================================================================================

    '''<summary>Get a list of all header elements.</summary>
    Private Function ReadHeader(ByRef BaseIn As System.IO.StreamReader) As List(Of sHeaderElement)

        Dim EndReached As Boolean = False
        Dim RetVal As New List(Of sHeaderElement)
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
                    Dim NewHeaderElement As sHeaderElement = Nothing
                    NewHeaderElement.Element = NewLine.Substring(0, 8).Trim
                    NewLine = NewLine.Substring(10)
                    Dim Splitted As String() = Split(NewLine, "/")
                    NewHeaderElement.Value = Splitted(0).Trim
                    If Splitted.Length > 1 Then NewHeaderElement.Comment = Splitted(1)
                    RetVal.Add(NewHeaderElement)
                    Select Case NewHeaderElement.Element
                        Case "BITPIX" : MyProps.BitPix = CInt(NewHeaderElement.Value)
                        Case "NAXIS1" : MyProps.Width = CInt(NewHeaderElement.Value)
                        Case "NAXIS2" : MyProps.Height = CInt(NewHeaderElement.Value)
                        Case "NAXIS3" : MyProps.ColorValues = CInt(NewHeaderElement.Value)
                        Case "BZERO" : MyProps.BZERO = Val(NewHeaderElement.Value.Replace(",", "."))
                        Case "BSCALE" : MyProps.BSCALE = Val(NewHeaderElement.Value.Replace(",", "."))
                    End Select
                End If
            Loop Until EndReached = True
        Loop Until EndReached = True

        MyProps.DataStartIdx = BlocksRead * HeaderBlockSize
        Return RetVal

    End Function

End Class
