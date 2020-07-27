Option Explicit On
Option Strict On

'http://lclevy.free.fr/nef/
' -> "II" means little-endian mean MSB right

Public Class cNEFReader

    Dim SensorX As Integer = 8256       'sensor width  [pixel]
    Dim SensorY As Integer = 5504       'sensor height [pixel]
    Dim Resolution As Integer = 14

    Public Sub Read(ByVal FileName As String)

        Dim BinIn As New System.IO.BinaryReader(System.IO.File.OpenRead(FileName))

        Dim Offset As Long = 5394944
        Dim Length As Int64 = (SensorX * Resolution)
        Length \= 8
        Dim ExtraByte As Integer = 5

        BinIn.BaseStream.Seek(Offset, IO.SeekOrigin.Begin)

        Dim Buffer() As Byte : ReDim Buffer(CInt(Length - 1) + ExtraByte)
        Dim X As BitArray

        'Dim Dbg As String = GetBits(X)
        'Dim HistPix As Dictionary(Of String, Integer) = DictHist(Dbg, Resolution)
        'HistPix = HistPix.SortDictionary

        Dim ImageData(SensorX - 1, SensorY - 1) As Integer
        For LineIdx As Integer = 0 To SensorY - 1
            BinIn.Read(Buffer, 0, Buffer.Length)
            X = New BitArray(Buffer)
            Dim BitPtr As Integer = 0
            Dim PixelPtr As Integer = 0
            For PixelIdx As Int64 = 0 To SensorX - 1
                Dim PixelVal As UInt16 = 0
                For BitIdx As Integer = 0 To Resolution - 1
                    Dim ToAdd As UInt16 = (CByte(X.Get(BitPtr)) << BitIdx)
                    PixelVal = PixelVal + ToAdd
                    BitPtr += 1
                Next BitIdx
                ImageData(PixelPtr, LineIdx) = PixelVal
                PixelPtr += 1
            Next PixelIdx
        Next LineIdx

        BinIn.Close()
        Dim FileToGenerate As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(FileName), "Converted.fits")
        cFITSWriter.Write(FileToGenerate, ImageData, cFITSWriter.eBitPix.Int32)
        Process.Start(FileToGenerate)

        MsgBox("OK")


    End Sub

    '''<summary>Get a string with bits from the passed BitArray.</summary>
    '''<param name="Buffer">Array of bits.</param>
    '''<returns>String with 0 and 1.</returns>
    Private Function GetBits(ByRef Buffer As BitArray) As String
        Dim RetVal As New List(Of String)
        Dim Cnt As Integer = 0
        For Each Entry As Boolean In Buffer
            RetVal.Add(CStr(IIf(Entry = True, "1", "0")))
            Cnt += 1
        Next Entry
        Return Join(RetVal.ToArray, String.Empty)
    End Function

    Private Function DictHist(ByRef Vector As String, ByVal Res As Integer) As Dictionary(Of String, Integer)
        Dim RetVal As New Dictionary(Of String, Integer)
        For Idx As Integer = 0 To Vector.Length - 1 Step Res
            Dim PixelVal As String = "00" & Vector.Substring(Idx, Res)
            Dim AsBigEnd As String = PixelVal.Substring(7, 8) & PixelVal.Substring(0, 8)
            PixelVal = ToHex(AsBigEnd)
            If RetVal.ContainsKey(PixelVal) = True Then
                RetVal(PixelVal) += 1
            Else
                RetVal.Add(PixelVal, 1)
            End If
        Next Idx
        Return RetVal
    End Function

    Private Function ToHex(ByRef ByteVector As String) As String
        Dim RetVal As New List(Of String)
        For Idx As Integer = 0 To ByteVector.Length - 1 Step 4
            Select Case ByteVector.Substring(Idx, 4)
                Case "0000" : RetVal.Add("0")
                Case "0001" : RetVal.Add("1")
                Case "0010" : RetVal.Add("2")
                Case "0011" : RetVal.Add("3")
                Case "0100" : RetVal.Add("4")
                Case "0101" : RetVal.Add("5")
                Case "0110" : RetVal.Add("6")
                Case "0111" : RetVal.Add("7")
                Case "1000" : RetVal.Add("8")
                Case "1001" : RetVal.Add("9")
                Case "1010" : RetVal.Add("A")
                Case "1011" : RetVal.Add("B")
                Case "1100" : RetVal.Add("C")
                Case "1101" : RetVal.Add("D")
                Case "1110" : RetVal.Add("E")
                Case "1111" : RetVal.Add("F")
                Case Else
                    MsgBox("!!!")
            End Select
        Next Idx
        Return Join(RetVal.ToArray, String.Empty)
    End Function

End Class