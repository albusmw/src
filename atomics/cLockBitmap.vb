Option Explicit On
Option Strict On

'''<summary>Class for fast access to bitmap data in memory using as 32-bit buffer.</summary>
Public Class cLockBitmap32Bit

    '''<summary>Pixel format to use - fixed to 32-bit.</summary>
    Private ReadOnly PixelFormat As System.Drawing.Imaging.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppArgb

    Public BitmapToProcess As Drawing.Bitmap = Nothing
    Public BitmapData As Drawing.Imaging.BitmapData = Nothing

    Private BitmapDataPtr As IntPtr = IntPtr.Zero

    Public Pixels As Int32()

    Public Property Width() As Integer = -1
    Public Property Height() As Integer = -1

    '''<summary>Init a new bitmap with the given width and height.</summary>
    Public Sub New(ByVal Width As Integer, ByVal Height As Integer)
        If Width > 0 And Height > 0 Then
            Me.BitmapToProcess = New Drawing.Bitmap(Width, Height, PixelFormat)
        End If
    End Sub

    '''<summary>Lock bitmap data.</summary>
    Public Sub LockBits()

        'Get width and height of bitmap
        Width = BitmapToProcess.Width
        Height = BitmapToProcess.Height

        'Check if bpp (Bits Per Pixel) is 32
        Select Case System.Drawing.Bitmap.GetPixelFormatSize(BitmapToProcess.PixelFormat)
            Case 32
                'Supported
            Case Else
                Throw New ArgumentException("Only 32 bpp images are supported.")
        End Select

        'Lock bitmap and return bitmap data
        BitmapData = BitmapToProcess.LockBits(New Drawing.Rectangle(0, 0, Width, Height), Drawing.Imaging.ImageLockMode.ReadWrite, BitmapToProcess.PixelFormat)

        'Create byte array to copy pixel values
        Pixels = New Int32((Width * Height) - 1) {}
        BitmapDataPtr = BitmapData.Scan0

        'Copy data from pointer to array
        Runtime.InteropServices.Marshal.Copy(BitmapDataPtr, Pixels, 0, Pixels.Length)

    End Sub

    '''<summary>Unlock bitmap data.</summary>
    Public Sub UnlockBits()
        Try
            Runtime.InteropServices.Marshal.Copy(Pixels, 0, BitmapDataPtr, Pixels.Length)       'Copy data from byte array to pointer
            BitmapToProcess.UnlockBits(BitmapData)                                              'Unlock bitmap data
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class

'''<summary>Class for fast access to bitmap data in memory.</summary>
Public Class cLockBitmap

    Public Enum RGBChannel
        R
        G
        B
    End Enum

    Private PixelFormat As System.Drawing.Imaging.PixelFormat = System.Drawing.Imaging.PixelFormat.Format24bppRgb

    Public Property ColorBytesPerPixel As Integer = 0


    Public BitmapToProcess As Drawing.Bitmap = Nothing
    Public BitmapData As Drawing.Imaging.BitmapData = Nothing

    Private BitmapDataPtr As IntPtr = IntPtr.Zero

    Public Pixels As Byte()

    Public Property Width() As Integer = -1
    Public Property Height() As Integer = -1
    Public Property InvalidColor As Drawing.Color = Drawing.Color.HotPink

    '''<summary>Init a new bitmap with the given width and height.</summary>
    Public Sub New(ByVal Width As Integer, ByVal Height As Integer)
        If Width > 0 And Height > 0 Then
            Me.BitmapToProcess = New Drawing.Bitmap(Width, Height, PixelFormat)
        End If
    End Sub

    '''<summary>Init a new bitmap with the given width and height.</summary>
    Public Sub New(ByVal Width As Integer, ByVal Height As Integer, ByVal PixelFormatToUse As System.Drawing.Imaging.PixelFormat)
        If Width > 0 And Height > 0 Then
            PixelFormat = PixelFormatToUse
            Me.BitmapToProcess = New Drawing.Bitmap(Width, Height, PixelFormat)
        End If
    End Sub

    '''<summary>Lock bitmap data.</summary>
    Public Sub LockBits()

        'Get width and height of bitmap
        Width = BitmapToProcess.Width
        Height = BitmapToProcess.Height

        'Get total locked pixels count
        Dim PixelCount As Integer = Width * Height

        'Create rectangle to lock
        Dim rect As New Drawing.Rectangle(0, 0, Width, Height)

        'Get source bitmap pixel format size
        ColorBytesPerPixel = System.Drawing.Bitmap.GetPixelFormatSize(BitmapToProcess.PixelFormat) \ 8

        'Check if bpp (Bits Per Pixel) is 8, 24, or 32
        Select Case ColorBytesPerPixel
            Case 1, 3, 4
                'Supported
            Case Else
                Throw New ArgumentException("Only 8, 24 and 32 bpp images are supported.")
        End Select

        'Lock bitmap and return bitmap data
        BitmapData = BitmapToProcess.LockBits(rect, Drawing.Imaging.ImageLockMode.ReadWrite, BitmapToProcess.PixelFormat)

        'Create byte array to copy pixel values
        Pixels = New Byte(BitmapData.Stride * Height - 1) {}
        BitmapDataPtr = BitmapData.Scan0

        'Copy data from pointer to array
        Runtime.InteropServices.Marshal.Copy(BitmapDataPtr, Pixels, 0, Pixels.Length)

    End Sub

    '''<summary>Unlock bitmap data.</summary>
    Public Sub UnlockBits()
        Try
            Runtime.InteropServices.Marshal.Copy(Pixels, 0, BitmapDataPtr, Pixels.Length)    'Copy data from byte array to pointer
            BitmapToProcess.UnlockBits(BitmapData)                                           'Unlock bitmap data
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '''<summary>Get the color of the specified pixel.</summary>
    '''<param name="X">X coordinate - from left to right, X=0 is top-left.</param>
    '''<param name="Y">Y coordinate - from top to bottom, Y=0 is top.</param>
    '''<returns>Color value.</returns>
    Public Function GetPixel(x As Integer, y As Integer) As Drawing.Color

        Dim ColorToSet As Drawing.Color = Drawing.Color.Empty

        ' Get start index of the specified pixel
        Dim i As Integer = (y * BitmapData.Stride) + (x * ColorBytesPerPixel)

        If i > Pixels.Length - ColorBytesPerPixel Then
            Throw New IndexOutOfRangeException()
        End If

        If ColorBytesPerPixel = 4 Then
            ' For 32 bpp get Red, Green, Blue and Alpha
            Dim b As Byte = Pixels(i)
            Dim g As Byte = Pixels(i + 1)
            Dim r As Byte = Pixels(i + 2)
            Dim a As Byte = Pixels(i + 3)
            ' a
            ColorToSet = Drawing.Color.FromArgb(a, r, g, b)
        End If

        If ColorBytesPerPixel = 3 Then
            ' For 24 bpp get Red, Green and Blue
            Dim b As Byte = Pixels(i)
            Dim g As Byte = Pixels(i + 1)
            Dim r As Byte = Pixels(i + 2)
            ColorToSet = Drawing.Color.FromArgb(r, g, b)
        End If

        If ColorBytesPerPixel = 1 Then
            ' For 8 bpp get color value (Red, Green and Blue values are the same)
            Dim c As Byte = Pixels(i)
            ColorToSet = Drawing.Color.FromArgb(c, c, c)
        End If

        Return ColorToSet

    End Function

    '''<summary>Set the complete image to the given color.</summary>
    Public Sub SetAll(ByVal Color As Drawing.Color)

        ' Get start index of the specified pixel
        Dim R As Byte = Color.R
        Dim G As Byte = Color.G
        Dim B As Byte = Color.B

        For IdxX As Integer = 0 To Me.Width - 1
            For IdxY As Integer = 0 To Me.Height - 1
                SetPixel(IdxX, IdxY, R, G, B)
            Next IdxY
        Next IdxX

    End Sub

    '''<summary>Set the color of the specified pixel</summary>
    '''<param name="X">X coordinate - from left to right, X=0 is top-left.</param>
    '''<param name="Y">Y coordinate - from top to bottom, Y=0 is top.</param>
    Public Sub SetPixel(x As Integer, y As Integer, ByVal Color As Drawing.Color)

        ' Get start index of the specified pixel
        Dim i As Integer = (y * BitmapData.Stride) + (x * ColorBytesPerPixel)
        If ColorBytesPerPixel = 3 Then
            SetPixel(i, Color.R, Color.G, Color.B)
        End If

    End Sub

    '''<summary>Set the color of the specified pixel</summary>
    Public Sub SetPixel(ByVal PixelIndex As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
        Pixels(PixelIndex) = B
        Pixels(PixelIndex + 1) = G
        Pixels(PixelIndex + 2) = R
    End Sub

    '''<summary>Set the color of the specified pixel</summary>
    '''<param name="X">X coordinate - from left to right, X=0 is top-left.</param>
    '''<param name="Y">Y coordinate - from top to bottom, Y=0 is top.</param>
    Public Sub SetPixel(x As Integer, y As Integer, ByVal R As Double, ByVal G As Double, ByVal B As Double)

        ' Get start index of the specified pixel
        Dim i As Integer = (y * BitmapData.Stride) + (x * ColorBytesPerPixel)

        'Cache the invalid color code
        Dim Invalid_R As Byte = InvalidColor.R
        Dim Invalid_G As Byte = InvalidColor.G
        Dim Invalid_B As Byte = InvalidColor.B

        If ColorBytesPerPixel = 3 Then
            ' For 24 bpp set Red, Green and Blue
            If R >= 0 And R <= 255 And G >= 0 And G <= 255 And B >= 0 And B <= 255 Then
                Pixels(i) = CByte(B)
                Pixels(i + 1) = CByte(G)
                Pixels(i + 2) = CByte(R)
            Else
                Pixels(i) = Invalid_B
                Pixels(i + 1) = Invalid_G
                Pixels(i + 2) = Invalid_R
            End If
        End If

    End Sub

    '''<summary>Set the color of the specified pixel using a "linear" grayscale maping.</summary>
    '''<param name="BaseOffset">Base offset within the data.</param>
    Public Sub SetPixel(ByVal BaseOffset As Integer, ByVal Value As Double)
        If ColorBytesPerPixel = 3 Then
            ' For 24 bpp set Red, Green and Blue
            If Value >= 0 And Value <= 255 Then
                Dim ByteValue As Byte = CByte(Value)
                Pixels(BaseOffset) = ByteValue
                Pixels(BaseOffset + 1) = ByteValue
                Pixels(BaseOffset + 2) = ByteValue
            Else
                Pixels(BaseOffset) = InvalidColor.B
                Pixels(BaseOffset + 1) = InvalidColor.G
                Pixels(BaseOffset + 2) = InvalidColor.R
            End If
        End If
    End Sub

    '''<summary>Set the color of the specified pixel using a "linear" grayscale maping.</summary>
    '''<param name="X">X coordinate - from left to right, X=0 is top-left.</param>
    '''<param name="Y">Y coordinate - from top to bottom, Y=0 is top.</param>
    Public Sub SetPixel(x As Integer, y As Integer, ByVal Value As Double)

        ' Get start index of the specified pixel
        Dim BaseOffset As Integer = (y * BitmapData.Stride) + (x * ColorBytesPerPixel)

        If ColorBytesPerPixel = 3 Then
            ' For 24 bpp set Red, Green and Blue
            If Value >= 0 And Value <= 255 Then
                Dim ByteValue As Byte = CByte(Value)
                Pixels(BaseOffset) = ByteValue
                Pixels(BaseOffset + 1) = ByteValue
                Pixels(BaseOffset + 2) = ByteValue
            Else
                Pixels(BaseOffset) = InvalidColor.B
                Pixels(BaseOffset + 1) = InvalidColor.G
                Pixels(BaseOffset + 2) = InvalidColor.R
            End If
        End If

    End Sub

    '''<summary>Calculate a linear grayscale image from the passed data.</summary>
    '''<param name="ImageData">Data to calculate cLockBitmap from.</param>
    '''<param name="MaxData">Data maximum.</param>
    Public Shared Function GetGrayscaleImage(ByRef ImageData(,) As UInt16, ByVal MaxData As Double) As cLockBitmap
        Dim OutputImage As New cLockBitmap(ImageData.GetUpperBound(0), ImageData.GetUpperBound(1))
        OutputImage.LockBits()
        Dim Stride As Integer = OutputImage.BitmapData.Stride
        Dim BytePerPixel As Integer = OutputImage.ColorBytesPerPixel
        Dim YOffset As Integer = 0
        For Y As Integer = 0 To OutputImage.Height - 1
            Dim BaseOffset As Integer = YOffset
            For X As Integer = 0 To OutputImage.Width - 1
                Dim DispVal As Integer = CInt(ImageData(X, Y) * (255 / MaxData))
                Dim Coloring As Drawing.Color = Drawing.Color.FromArgb(DispVal, DispVal, DispVal)
                OutputImage.Pixels(BaseOffset) = Coloring.B
                OutputImage.Pixels(BaseOffset + 1) = Coloring.G
                OutputImage.Pixels(BaseOffset + 2) = Coloring.R
                BaseOffset += BytePerPixel
            Next X
            YOffset += Stride
        Next Y
        OutputImage.UnlockBits()
        Return OutputImage
    End Function

End Class