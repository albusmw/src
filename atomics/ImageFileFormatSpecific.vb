Option Explicit On
Option Strict On

'''<summary>Load and store special file formats.</summary>
Public Class ImageFileFormatSpecific

    Public Shared Sub SaveTIFF_Format16bppGrayScale(ByVal FileName As String, ByRef Data(,) As UInt16)

        'https://bitmiracle.github.io/libtiff.net/help/articles/KB/grayscale-color.html

        Using output As BitMiracle.LibTiff.Classic.Tiff = BitMiracle.LibTiff.Classic.Tiff.Open(FileName, "w")

            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGEWIDTH, Data.GetUpperBound(0) + 1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGELENGTH, Data.GetUpperBound(1) + 1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.SAMPLESPERPIXEL, 1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.BITSPERSAMPLE, 16)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.ORIENTATION, BitMiracle.LibTiff.Classic.Orientation.TOPLEFT)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.ROWSPERSTRIP, Data.GetUpperBound(1) + 1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.XRESOLUTION, 88.0)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.YRESOLUTION, 88.0)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.RESOLUTIONUNIT, BitMiracle.LibTiff.Classic.ResUnit.INCH)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.PLANARCONFIG, BitMiracle.LibTiff.Classic.PlanarConfig.CONTIG)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.PHOTOMETRIC, BitMiracle.LibTiff.Classic.Photometric.MINISBLACK)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.COMPRESSION, BitMiracle.LibTiff.Classic.Compression.NONE)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.FILLORDER, BitMiracle.LibTiff.Classic.FillOrder.MSB2LSB)

            Dim BitWidth As Integer = 2
            For i As Integer = 0 To Data.GetUpperBound(1)
                Dim Buffer((BitWidth * (Data.GetUpperBound(0) + 1)) - 1) As Byte
                Dim BufferPtr As Integer = 0
                For j As Integer = 0 To Data.GetUpperBound(0)
                    Dim Pattern() As Byte = BitConverter.GetBytes(Data(j, i))
                    Buffer(BufferPtr) = Pattern(0)
                    Buffer(BufferPtr + 1) = Pattern(1)
                    BufferPtr += 2
                Next j
                output.WriteScanline(Buffer, i)
            Next i

            output.WriteDirectory()

        End Using

    End Sub

    Public Shared Sub SaveTIFF_Format48bppColor(ByVal FileName As String, ByRef Data() As cStatMultiThread_UInt16.sImgData_UInt16)

        'https://bitmiracle.github.io/libtiff.net/help/articles/KB/grayscale-color.html

        Using output As BitMiracle.LibTiff.Classic.Tiff = BitMiracle.LibTiff.Classic.Tiff.Open(FileName, "w")

            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGEWIDTH, Data(0).NAXIS1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGELENGTH, Data(0).NAXIS2)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.SAMPLESPERPIXEL, 3)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.BITSPERSAMPLE, 16)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.ORIENTATION, BitMiracle.LibTiff.Classic.Orientation.TOPLEFT)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.ROWSPERSTRIP, Data(0).NAXIS1)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.XRESOLUTION, 88.0)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.YRESOLUTION, 88.0)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.RESOLUTIONUNIT, BitMiracle.LibTiff.Classic.ResUnit.INCH)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.PLANARCONFIG, BitMiracle.LibTiff.Classic.PlanarConfig.CONTIG)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.PHOTOMETRIC, BitMiracle.LibTiff.Classic.Photometric.RGB)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.COMPRESSION, BitMiracle.LibTiff.Classic.Compression.NONE)
            output.SetField(BitMiracle.LibTiff.Classic.TiffTag.FILLORDER, BitMiracle.LibTiff.Classic.FillOrder.MSB2LSB)

            Dim BitWidth As Integer = 6
            For i As Integer = 0 To Data(0).Data.GetUpperBound(1)
                Dim Buffer((BitWidth * (Data(0).Data.GetUpperBound(0) + 1)) - 1) As Byte
                Dim BufferPtr As Integer = 0
                For j As Integer = 0 To Data(0).Data.GetUpperBound(0)
                    Dim PatternR() As Byte = BitConverter.GetBytes(Data(0).Data(j, i))
                    Buffer(BufferPtr) = PatternR(0)
                    Buffer(BufferPtr + 1) = PatternR(1)
                    BufferPtr += 2
                    Dim PatternG() As Byte = BitConverter.GetBytes(Data(1).Data(j, i))
                    Buffer(BufferPtr) = PatternG(0)
                    Buffer(BufferPtr + 1) = PatternG(1)
                    BufferPtr += 2
                    Dim PatternB() As Byte = BitConverter.GetBytes(Data(2).Data(j, i))
                    Buffer(BufferPtr) = PatternB(0)
                    Buffer(BufferPtr + 1) = PatternB(1)
                    BufferPtr += 2
                Next j
                output.WriteScanline(Buffer, i)
            Next i

            output.WriteDirectory()

        End Using

    End Sub

    '''<summary>Generate a test image with 16Bit per color channel.</summary>
    '''<param name="FileName">File to generate.</param>
    Public Shared Sub SavePNG_Format48bppRGB(ByVal FileName As String, ByRef Data As List(Of UInt16(,)))

        Dim b16bpp As New Bitmap(Data(0).GetUpperBound(0) + 1, Data(0).GetUpperBound(1) + 1, System.Drawing.Imaging.PixelFormat.Format48bppRgb)
        Dim BytePerPixel As Integer = 6

        Dim rect As New Rectangle(0, 0, b16bpp.Width, b16bpp.Height)

        Dim bitmapData As System.Drawing.Imaging.BitmapData = b16bpp.LockBits(rect, Imaging.ImageLockMode.WriteOnly, b16bpp.PixelFormat)

        'Calculate the number of bytes required And allocate them.
        Dim bitmapBytes((b16bpp.Width * b16bpp.Height * BytePerPixel) - 1) As Byte

        'Fill the bitmap bytes with random data.
        For x As Integer = 0 To b16bpp.Width - 1
            For y As Integer = 0 To b16bpp.Height - 1
                Dim PixelIdx As Integer = (y * b16bpp.Width * BytePerPixel) + (x * BytePerPixel)
                Dim BitPatternR As Byte() = BitConverter.GetBytes(Data(0)(x, y))
                Dim BitPatternG As Byte() = BitConverter.GetBytes(Data(1)(x, y))
                Dim BitPatternB As Byte() = BitConverter.GetBytes(Data(2)(x, y))
                bitmapBytes(PixelIdx + 0) = BitPatternR(0)
                bitmapBytes(PixelIdx + 1) = BitPatternR(1)
                bitmapBytes(PixelIdx + 2) = BitPatternG(0)
                bitmapBytes(PixelIdx + 3) = BitPatternG(1)
                bitmapBytes(PixelIdx + 4) = BitPatternB(0)
                bitmapBytes(PixelIdx + 5) = BitPatternB(1)
            Next y
        Next x

        'Copy the randomized bits to the bitmap pointer.
        Dim Pointer As IntPtr = bitmapData.Scan0
        Runtime.InteropServices.Marshal.Copy(bitmapBytes, 0, Pointer, bitmapBytes.Length)

        'Unlock the bitmap, we're all done.
        b16bpp.UnlockBits(bitmapData)

        b16bpp.Save(FileName, Imaging.ImageFormat.Png)

    End Sub

    '''<summary>Generate a test image with 16Bit per color channel.</summary>
    '''<param name="FileName">File to generate.</param>
    Public Shared Sub Test48BPP(ByVal FileName As String)

        Dim b16bpp As New Bitmap(5000, 5000, System.Drawing.Imaging.PixelFormat.Format48bppRgb)
        Dim BytePerPixel As Integer = 6

        Dim rect As New Rectangle(0, 0, b16bpp.Width, b16bpp.Height)

        Dim bitmapData As System.Drawing.Imaging.BitmapData = b16bpp.LockBits(rect, Imaging.ImageLockMode.WriteOnly, b16bpp.PixelFormat)

        'Calculate the number of bytes required And allocate them.
        Dim bitmapBytes((b16bpp.Width * b16bpp.Height * BytePerPixel) - 1) As Byte

        'Fill the bitmap bytes with random data.
        Dim Rnd As New Random
        Dim RunIdx As Long = 0
        For x As Integer = 0 To b16bpp.Width - 1
            For y As Integer = 0 To b16bpp.Height - 1
                Dim PixelIdx As Integer = (y * b16bpp.Width * BytePerPixel) + (x * BytePerPixel)
                Dim BitPatternR As Byte() = BitConverter.GetBytes(CType(Rnd.Next(0, UInt16.MaxValue + 1), UInt16))
                Dim BitPatternG As Byte() = BitConverter.GetBytes(CType(Rnd.Next(0, UInt16.MaxValue + 1), UInt16))
                Dim BitPatternB As Byte() = BitConverter.GetBytes(CType(Rnd.Next(0, UInt16.MaxValue + 1), UInt16))
                bitmapBytes(PixelIdx + 0) = BitPatternR(0)
                bitmapBytes(PixelIdx + 1) = BitPatternR(1)
                bitmapBytes(PixelIdx + 2) = BitPatternG(0)
                bitmapBytes(PixelIdx + 3) = BitPatternG(1)
                bitmapBytes(PixelIdx + 4) = BitPatternB(0)
                bitmapBytes(PixelIdx + 5) = BitPatternB(1)
                RunIdx += 1
            Next y
        Next x

        'Copy the randomized bits to the bitmap pointer.
        Dim Pointer As IntPtr = bitmapData.Scan0
        Runtime.InteropServices.Marshal.Copy(bitmapBytes, 0, Pointer, bitmapBytes.Length)

        'Unlock the bitmap, we're all done.
        b16bpp.UnlockBits(bitmapData)

        b16bpp.Save(FileName, Imaging.ImageFormat.Png)

    End Sub

    Public Shared Function Make1bpp(ByVal bmpIN As Bitmap) As Bitmap

        Dim bmpOUT As Bitmap
        bmpOUT = New Bitmap(bmpIN.Width, bmpIN.Height, System.Drawing.Imaging.PixelFormat.Format1bppIndexed)
        bmpOUT.SetResolution(bmpIN.HorizontalResolution, bmpIN.VerticalResolution)

        ' seems like I've got this crap in this program about 100x.
        If bmpIN.PixelFormat <> System.Drawing.Imaging.PixelFormat.Format16bppRgb555 Then
            Throw New ApplicationException("hand-coded routine can only understand image format of Format16bppRgb555 but this image is " &
              bmpIN.PixelFormat.ToString & ". Either change the format or code this sub to handle that format, too.")
        End If

        ' lock image bytes
        Dim bmdIN As System.Drawing.Imaging.BitmapData = bmpIN.LockBits(New Rectangle(0, 0, bmpIN.Width, bmpIN.Height),
            Imaging.ImageLockMode.ReadWrite, bmpIN.PixelFormat)
        ' lock image bytes
        Dim bmdOUT As System.Drawing.Imaging.BitmapData = bmpOUT.LockBits(New Rectangle(0, 0, bmpOUT.Width, bmpOUT.Height),
            Imaging.ImageLockMode.ReadWrite, bmpOUT.PixelFormat)

        ' Allocate room for the data.
        Dim bytesIN(bmdIN.Stride * bmdIN.Height) As Byte
        Dim bytesOUT(bmdOUT.Stride * bmdOUT.Height) As Byte

        'Copy the data into the PixBytes array. 
        Runtime.InteropServices.Marshal.Copy(bmdIN.Scan0, bytesIN, 0, CInt(bmdIN.Stride * bmpIN.Height))
        ' > this val = white pix. (each of the 3 pix in the rgb555 can hold 32 levels... 2^5 huh.)
        Dim bThresh As Byte = CByte((32 * 3) * 0.66)
        ' transfer the pixels
        For y As Integer = 0 To bmpIN.Height - 1
            Dim outpos As Integer = y * bmdOUT.Stride
            Dim instart As Integer = y * bmdIN.Stride
            Dim byteval As Byte = 0
            Dim bitpos As Byte = 128
            Dim pixval As Integer
            Dim pixgraylevel As Integer
            For inpos As Integer = instart To instart + bmdIN.Stride - 1 Step 2
                pixval = 256 * bytesIN(inpos + 1) + bytesIN(inpos) ' DEPENDANT ON Format16bppRgb555
                pixgraylevel = ((pixval) And 31) + ((pixval >> 5) And 31) + ((pixval >> 10) And 31)
                If pixgraylevel > bThresh Then ' DEPENDANT ON Format16bppRgb555
                    byteval = byteval Or bitpos
                End If
                bitpos = bitpos >> 1
                If bitpos = 0 Then
                    bytesOUT(outpos) = byteval
                    byteval = 0
                    bitpos = 128
                    outpos += 1
                End If
            Next
            If bitpos <> 0 Then ' stick a fork in any unfinished busines.
                bytesOUT(outpos) = byteval
            End If
        Next
        ' unlock image bytes
        ' Copy the data back into the bitmap. 
        Runtime.InteropServices.Marshal.Copy(bytesOUT, 0, bmdOUT.Scan0, bmdOUT.Stride * bmdOUT.Height)
        ' Unlock the bitmap.
        bmpIN.UnlockBits(bmdIN)
        bmpOUT.UnlockBits(bmdOUT)
        ' futile attempt to free memory.
        ReDim bytesIN(0)
        ReDim bytesOUT(0)
        ' return new bmp.
        Return bmpOUT

    End Function

    '''<summary>Return the filters for the save dialog.</summary>
    Public Shared Function SaveImageFileFilters() As String
        Dim RetVal As New List(Of String)
        RetVal.Add("JPEG file (*.jpg)|*.jpg")
        RetVal.Add("PNG file (*.png)|*.png")
        Return Join(RetVal.ToArray, "|")
    End Function

    '''<summary>Save the content in the given format specified by the file name.</summary>
    Public Shared Sub SaveImageFile(ByVal FileName As String, ByRef Content As Bitmap)

        Dim Param As New Imaging.EncoderParameters
        Dim Extension As String = System.IO.Path.GetExtension(FileName).ToUpper
        Select Case Extension
            Case ".JPG", ".JPEG"
                Dim Encoder As Imaging.ImageCodecInfo = GetEncoder(Imaging.ImageFormat.Jpeg)
                Dim EncoderParameters As New Imaging.EncoderParameters
                EncoderParameters.Param(0) = New Imaging.EncoderParameter(Imaging.Encoder.Quality, CLng(InputBox("Quality [0..100]", "Quality", "90")))
                Content.Save(FileName, Encoder, EncoderParameters)
            Case ".PNG"
                Dim Encoder As Imaging.ImageCodecInfo = GetEncoder(Imaging.ImageFormat.Png)
                Dim EncoderParameters As New Imaging.EncoderParameters
                EncoderParameters.Param(0) = New Imaging.EncoderParameter(Imaging.Encoder.Quality, CLng(InputBox("Quality [0..100]", "Quality", "90")))
                EncoderParameters.Param(1) = New Imaging.EncoderParameter(Imaging.Encoder.Compression, Imaging.EncoderValue.CompressionLZW)
                Content.Save(FileName, Encoder, EncoderParameters)
        End Select

    End Sub

    '''<summary>Get the base codec info for the requested format.</summary>
    '''<param name="RequestedFormat">Format requested.</param>
    '''<returns>Codec or nothing if codes is not found.</returns>
    Private Shared Function GetEncoder(ByVal RequestedFormat As Imaging.ImageFormat) As Imaging.ImageCodecInfo
        Dim codes As Imaging.ImageCodecInfo() = Imaging.ImageCodecInfo.GetImageDecoders
        For Each AvailableCodec As Imaging.ImageCodecInfo In codes
            If AvailableCodec.FormatID = RequestedFormat.Guid Then Return AvailableCodec
        Next AvailableCodec
        Return Nothing
    End Function


End Class



'Dim stream As New IO.FileStream(sfdMain.FileName, IO.FileMode.Create)
'Dim encoder As New System.Windows.Media.Imaging.TiffBitmapEncoder()
'encoder.Compression = Windows.Media.Imaging.TiffCompressOption.Zip
'encoder.Frames.Add(Windows.Media.Imaging.BitmapFrame.Create(New System.IO.MemoryStream(cLockBitmap.CalculateOutputBitmap(.ImageData, LastStat.MonoStatistics.Max.Key).Pixels)))
'encoder.Save(stream)