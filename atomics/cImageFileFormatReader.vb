Option Explicit
Option Strict On

Public Class cImageFileFormatReader

    ''' <summary>Use DCRAW.exe to convert a camera raw to a portable format.</summary>
    ''' <returns>Name of the converted file or empty string in case of an error.</returns>
    Public Shared Function UseDCRaw(ByVal DCRawEXE As String, ByVal FileName As String, ByRef Output As String()) As String

        Dim StartInfo As New ProcessStartInfo
        StartInfo.FileName = DCRawEXE
        StartInfo.Arguments = "-v -6 -W -g 1 1 -D " & Chr(34) & FileName & Chr(34)
        StartInfo.UseShellExecute = False
        StartInfo.CreateNoWindow = True
        StartInfo.RedirectStandardOutput = True
        StartInfo.RedirectStandardError = True
        Dim DCRaw As New Process
        DCRaw.StartInfo = StartInfo
        DCRaw.Start()
        DCRaw.WaitForExit()
        Output = DCRaw.StandardError.ReadToEnd.Trim.Split(Chr(10))

        'Check the output for the fle name generated
        Dim OutKeyWords As String = "Writing data to"
        For Each Line As String In Output
            If Line.StartsWith(OutKeyWords) Then
                Line = Line.Replace(OutKeyWords, String.Empty).Replace("...", String.Empty).TrimEnd
                If System.IO.File.Exists(Line) Then Return Line
            End If
        Next Line

        'Error occured ...
        Return String.Empty

    End Function

    ''' <summary>Read a TIFF file (16-bit).</summary>
    ''' <remarks>Requires a link to PresentationCore.</remarks>
    Public Shared Sub LoadTIFFData(ByVal FileName As String, ByRef ImageData(,) As Double)

        'Special TIFF loader
        Dim TifDec As New System.Windows.Media.Imaging.TiffBitmapDecoder(New Uri(FileName), Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat, Windows.Media.Imaging.BitmapCacheOption.OnLoad)
        Dim BmpFrame As System.Windows.Media.Imaging.BitmapFrame = TifDec.Frames(0)
        Dim Data(BmpFrame.PixelWidth * BmpFrame.PixelHeight - 1) As UShort
        BmpFrame.CopyPixels(Data, 2 * BmpFrame.PixelWidth, 0)

        ReDim ImageData(BmpFrame.PixelWidth - 1, BmpFrame.PixelHeight - 1)

        'Extract the "brightness" channel
        Dim ReadPtr As Integer = 0
        For Idx1 As Integer = 0 To BmpFrame.PixelHeight - 1
            For Idx2 As Integer = 0 To BmpFrame.PixelWidth - 1
                ImageData(Idx2, Idx1) = Data(ReadPtr)
                ReadPtr += 1
            Next Idx2
        Next Idx1

    End Sub


    ''' <summary>Read a Portable Anymap file (PNM / PBM / PGM / PPM).</summary>
    ''' <param name="FileName">File name to read in.</param>
    ''' <param name="ImageData">Data array to fill.</param>
    ''' <param name="ColorIdx">0-based color index in case of PixMap instead of GreyMap.</param>
    ''' <returns>TRUE if read was OK, FALSE else.</returns>
    ''' <remarks>Currently only "Portable Graymap Binary" is supported.</remarks>
    Public Shared Function LoadPortableAnyMapNew(ByRef FileName As String, ByRef ImageData(,) As Double, ByVal ColorIdx As Integer) As Boolean

        Dim RetVal As Boolean = False
        Dim File As New IO.FileStream(FileName, IO.FileMode.Open, IO.FileAccess.Read)
        Dim Header As New IO.BinaryReader(File)
        Dim Whitespaces As New List(Of Byte)({&H20, &H9, &HD, &HA})

        Dim ASCIIMode As Boolean = False
        Dim BitPerValue As Integer = 0

        'First byte must be a "P"
        If Header.ReadChar <> "P" Then Return False

        'The next ASCII number indicates the format
        Dim MagicNumber As Integer = CInt(Header.ReadChar.ToString)
        Select Case MagicNumber
            Case 1
                'Portable Bitmap  ASCII
                ASCIIMode = True : BitPerValue = 1
            Case 2
                'Portable Graymap ASCII
                ASCIIMode = True : BitPerValue = 8      '-> may get 16 if indicated with values > 255
            Case 3
                'Portable Pixmap  ASCII
                ASCIIMode = True : BitPerValue = 24      '-> may get 48 if indicated with values > 255
            Case 4
                'Portable Bitmap  Binary
                ASCIIMode = False : BitPerValue = 1
            Case 5
                'Portable Graymap Binary
                ASCIIMode = False : BitPerValue = 8      '-> may get 16 if indicated with values > 255
            Case 6
                'Portable Pixmap  Binary
                ASCIIMode = False : BitPerValue = 24      '-> may get 48 if indicated with values > 255
            Case 7
                'Portable Anymap
                ASCIIMode = False
            Case Else
                'Invalid
                Return False
        End Select

        'Read width (ASCII coded)
        Dim Width As String = String.Empty
        Do
            Dim ReadIn As Byte = Header.ReadByte
            If Whitespaces.Contains(ReadIn) = True Then
                If Width.Length > 0 Then Exit Do
            Else
                Width &= System.Text.ASCIIEncoding.ASCII.GetString(New Byte() {ReadIn})
            End If
        Loop Until File.Position = File.Length - 1
        Dim ArrayWidth As Integer = CInt(Width)

        'Read height (ASCII coded)
        Dim Height As String = String.Empty
        Do
            Dim ReadIn As Byte = Header.ReadByte
            If Whitespaces.Contains(ReadIn) = True Then
                If Height.Length > 0 Then Exit Do
            Else
                Height &= System.Text.ASCIIEncoding.ASCII.GetString(New Byte() {ReadIn})
            End If
        Loop Until File.Position = File.Length - 1
        Dim ArrayHeight As Integer = CInt(Height)

        'Read maximum value
        Select Case MagicNumber
            Case 1, 4
                'Not available
            Case 2, 3, 5, 6
                Dim MaxValueText As String = String.Empty
                Do
                    Dim ReadIn As Byte = Header.ReadByte
                    If Whitespaces.Contains(ReadIn) = True Then
                        If MaxValueText.Length > 0 Then Exit Do
                    Else
                        MaxValueText &= System.Text.ASCIIEncoding.ASCII.GetString(New Byte() {ReadIn})
                    End If
                Loop Until File.Position = File.Length - 1
                If CInt(MaxValueText) > 255 Then BitPerValue *= 2
        End Select

        'Move to data start (backwards ...)
        Dim ImageBytes As Long = (ArrayWidth * ArrayHeight) * (BitPerValue \ 8)
        File.Seek(File.Length - ImageBytes, IO.SeekOrigin.Begin)

        'Do
        '    Dim ReadIn As Byte = Header.ReadByte
        '    If Whitespaces.Contains(ReadIn) = False Then Exit Do
        'Loop Until File.Position = File.Length - 1

        'Read
        ReDim ImageData(ArrayWidth - 1, ArrayHeight - 1)
        Select Case MagicNumber
            Case 5
                Select Case BitPerValue
                    Case 8
                        For BayerPosY As Integer = 0 To ArrayHeight - 1
                            For BayerPosX As Integer = 0 To ArrayWidth - 1
                                ImageData(BayerPosX, BayerPosY) = Header.ReadByte
                            Next BayerPosX
                        Next BayerPosY
                        RetVal = True
                    Case 16
                        For BayerPosY As Integer = 0 To ArrayHeight - 1
                            For BayerPosX As Integer = 0 To ArrayWidth - 1
                                'Reading UInt16 does not read in the correct byte order ...
                                Dim Byte1 As Byte = Header.ReadByte
                                Dim Byte2 As Byte = Header.ReadByte
                                ImageData(BayerPosX, BayerPosY) = BitConverter.ToUInt16(New Byte() {Byte2, Byte1}, 0)
                            Next BayerPosX
                        Next BayerPosY
                        RetVal = True
                    Case Else
                        RetVal = False
                End Select
            Case 6
                Dim Matrix(2) As Double
                Select Case BitPerValue
                    Case 24
                        For BayerPosY As Integer = 0 To ArrayHeight - 1
                            For BayerPosX As Integer = 0 To ArrayWidth - 1
                                Matrix(0) = Header.ReadByte
                                Matrix(1) = Header.ReadByte
                                Matrix(2) = Header.ReadByte
                                ImageData(BayerPosX, BayerPosY) = Matrix(ColorIdx)
                            Next BayerPosX
                        Next BayerPosY
                        RetVal = True
                    Case 48
                        For BayerPosY As Integer = 0 To ArrayHeight - 1
                            For BayerPosX As Integer = 0 To ArrayWidth - 1
                                'Reading UInt16 does not read in the correct byte order ...
                                Dim Byte1 As Byte = Header.ReadByte
                                Dim Byte2 As Byte = Header.ReadByte
                                Matrix(0) = BitConverter.ToUInt16(New Byte() {Byte2, Byte1}, 0)
                                Dim Byte3 As Byte = Header.ReadByte
                                Dim Byte4 As Byte = Header.ReadByte
                                Matrix(1) = BitConverter.ToUInt16(New Byte() {Byte4, Byte3}, 0)
                                Dim Byte5 As Byte = Header.ReadByte
                                Dim Byte6 As Byte = Header.ReadByte
                                Matrix(2) = BitConverter.ToUInt16(New Byte() {Byte6, Byte5}, 0)
                                ImageData(BayerPosX, BayerPosY) = Matrix(ColorIdx)
                            Next BayerPosX
                        Next BayerPosY
                        RetVal = True
                    Case Else
                        RetVal = False
                End Select
            Case Else
                RetVal = False
        End Select

        Header.Close()

        Return RetVal

    End Function

End Class