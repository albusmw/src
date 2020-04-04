Option Explicit On
Option Strict On

Public Class ImageFileFormatSpecific

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

    Public Shared Sub Test16BPP(ByVal FileName As String)

        'Will generate a GDI+ error ...

        Dim b16bpp As New Bitmap(5000, 5000, System.Drawing.Imaging.PixelFormat.Format16bppGrayScale)
        Dim BytePerPixel As Integer = 2

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
                bitmapBytes(PixelIdx + 0) = BitPatternR(0)
                bitmapBytes(PixelIdx + 1) = BitPatternR(1)
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

End Class