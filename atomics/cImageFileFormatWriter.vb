Option Explicit
Option Strict On

Public Class cImageFileFormatWriter

    Public Shared Function SaveTIFFFile(ByRef FileName As String, ByRef ImageData(,) As Double, ByVal ColorIdx As Integer) As Boolean

        Dim Width As Integer = ImageData.GetUpperBound(0) + 1
        Dim Height As Integer = ImageData.GetUpperBound(1) + 1

        Dim Writeable As New System.Windows.Media.Imaging.WriteableBitmap(Width, Height, 88.0, 88.0, System.Windows.Media.PixelFormats.Gray16, Nothing)
        Dim Pixel((Width * Height) - 1) As UShort
        For Idx1 As Integer = 0 To Height - 1
            For Idx2 As Integer = 0 To Width - 1
                Pixel((Idx1 * Width) + Idx2) = CUShort(ImageData(Idx2, Idx1))
            Next Idx2
        Next Idx1

        Writeable.WritePixels(New Windows.Int32Rect(0, 0, Width, Height), Pixel, Width * 2, 0)

        Dim Encoder As New Windows.Media.Imaging.TiffBitmapEncoder
        Encoder.Frames.Add(Windows.Media.Imaging.BitmapFrame.Create(Writeable))
        Encoder.Save(System.IO.File.Create(FileName))

        Return True

    End Function

End Class