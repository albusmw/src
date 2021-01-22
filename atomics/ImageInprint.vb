Option Explicit On
Option Strict On

''' <summary>Add text / annotations / ... to a picture</summary>
Public Class ImageInprint

    ''' <summary>Parameters for the inprint.</summary>
    Public Structure sInprintParams
        ''' <summary>Text that should be printed.</summary>
        Public TextToPrint As String
        ''' <summary>Font size</summary>
        Public TextSize As Integer
        ''' <summary>Pixel value to set.</summary>
        Public PixelValue As Byte
        ''' <summary>Pixel value to set.</summary>
        Public PrintRight As Boolean
    End Structure

    ''' <summary>Add text to the passed bitmap data (8bit data).</summary>
    ''' <param name="BitmapValues">Bitmap data to manipulate.</param>
    Public Shared Sub Mono8BitText(ByRef BitmapToCreate As Bitmap, ByRef BitmapValues() As Byte, ByVal InprintParams As sInprintParams)

        Dim TestAsBitmap As Bitmap = DrawText(InprintParams.TextToPrint, InprintParams.TextSize)
        Dim NoText As Color = Color.FromArgb(0, 0, 0, 0)
        Dim NoTextColor As Byte = 0
        Dim StartX As Integer = 0 : If InprintParams.PrintRight = True Then StartX = BitmapToCreate.Width - TestAsBitmap.Width - 1
        Dim StartY As Integer = 0
        Dim IP1_X As Integer = 0
        For ScanX As Integer = StartX To StartX + TestAsBitmap.Width - 1
            Dim IP1_Y As Integer = 0
            For ScanY As Integer = StartY To StartY + TestAsBitmap.Height - 1
                If TestAsBitmap.GetPixel(IP1_X, IP1_Y) <> NoText Then
                    BitmapValues(ScanX + (ScanY * BitmapToCreate.Width)) = InprintParams.PixelValue     'pixel contains text color -> set in image
                Else
                    BitmapValues(ScanX + (ScanY * BitmapToCreate.Width)) = NoTextColor                  'pixel contains no text color -> make black (to get contrast)
                End If
                IP1_Y += 1
            Next ScanY
            IP1_X += 1
        Next ScanX

    End Sub

    Private Shared Function DrawText(ByVal TextToPrint As String, ByVal TextSize As Integer) As Bitmap

        'Create a text field and add it pixel-by-pixel
        Dim MyFont As New Font("Courier New", TextSize)
        Dim NoText As New Pen(Color.Black)
        Dim RequiredSize As Drawing.Size = TextRenderer.MeasureText(TextToPrint, MyFont)
        Dim Bmp As New Bitmap(RequiredSize.Width, RequiredSize.Height)
        Dim gra As Graphics = Graphics.FromImage(Bmp)
        gra.DrawRectangle(NoText, 0, 0, RequiredSize.Width, RequiredSize.Height)
        gra.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        gra.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixel
        gra.Clear(Color.Transparent)
        TextRenderer.DrawText(gra, TextToPrint, MyFont, New Point(0, 0), Color.Red)
        gra.Dispose()
        Return Bmp

    End Function

End Class