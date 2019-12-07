Option Explicit On
Option Strict On

'''<summary>Class to display passed data as an image.</summary>
Public Class cImageDisplay

    '''<summary>Get an color-mapped bitmap image from the given image data.</summary>
    Public Shared Function CalculateImageFromData(ByRef ImageData(,) As UInt32) As cLockBitmap

        Dim Width As Integer = ImageData.GetUpperBound(0) + 1
        Dim Height As Integer = ImageData.GetUpperBound(1) + 1

        Dim HistCalc As New cStatistics : HistCalc.Calculate(ImageData)

        'Generate output image
        Dim OutputImage As New cLockBitmap(Width, Height)

        Dim Invalid_R As Byte = OutputImage.InvalidColor.R
        Dim Invalid_G As Byte = OutputImage.InvalidColor.G
        Dim Invalid_B As Byte = OutputImage.InvalidColor.B

        'Auto-strech the LUT to the min and max in the image
        ' Color maps take values between 0 and 255
        Dim Min As Double = -1000
        Dim Max As Double = 1000
        Dim LinOffset As Double = Min - HistCalc.Minimum
        Dim LinScale As Double = (Max - Min) / (HistCalc.Maximum - HistCalc.Minimum)

        'Build a LUT for all colors present in the picture
        Dim LUT As New Dictionary(Of Double, Color)
        Dim LUTMin As Double = Double.MaxValue
        Dim LUTMax As Double = Double.MinValue
        For Each Entry As Double In HistCalc.Histogram.Keys
            Dim Scaled As Double = (Entry * LinScale) - LinOffset
            If Scaled > LUTMax Then LUTMax = Scaled
            If Scaled < LUTMin Then LUTMin = Scaled
            LUT.Add(Entry, cColorMaps.FalseColor(Scaled))
        Next Entry

        'Create the image
        OutputImage.LockBits()

        Dim Stride As Integer = OutputImage.BitmapData.Stride
        Dim BytePerPixel As Integer = OutputImage.ColorBytesPerPixel
        Dim YOffset As Integer = 0
        For Y As Integer = 0 To Height - 1
            Dim BaseOffset As Integer = YOffset
            For X As Integer = 0 To Width - 1
                Dim Coloring As Color = LUT(ImageData(X, Y))
                OutputImage.Pixels(BaseOffset) = Coloring.R
                OutputImage.Pixels(BaseOffset + 1) = Coloring.G
                OutputImage.Pixels(BaseOffset + 2) = Coloring.B
                BaseOffset += BytePerPixel
            Next X
            YOffset += Stride
        Next Y

        OutputImage.UnlockBits()

        Return OutputImage

    End Function

End Class