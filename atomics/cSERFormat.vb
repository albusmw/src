Option Explicit On
Option Strict On

'''<summary>Format helper and converter for the SER format handling.</summary>
'''<see cref="https://free-astro.org/index.php/SER"/>
Public Class cSERFormat

    Public Enum eColorID As Int32
        MONO = 0
        BAYER_RGGB = 8
        BAYER_GRBG = 9
        BAYER_GBRG = 10
        BAYER_BGGR = 11
        BAYER_CYYM = 16
        BAYER_YCMY = 17
        BAYER_YMCY = 18
        BAYER_MYYC = 19
    End Enum

    '''<summary>Structure holding the SER header information.</summary>
    Public Class cSERHeader

        '''<summary>Header size.</summary>
        Public Const SERHeaderLength As Integer = 178

        '''<summary>"LUCAM-RECORDER".</summary>
        Public Header As String = String.Empty
        '''<summary>Lumenera camera series ID - unused, default is 0.</summary>
        Public LuID As Int32 = Int32.MinValue
        '''<summary>Color ID.</summary>
        Public ColorID As eColorID
        '''<summary>0 for big-endian, 1 for little-endian byte order in 16 bit image data.</summary>
        Public LittleEndian As Int32 = Int32.MinValue
        '''<summary>Width of every image in pixel.</summary>
        Public FrameWidth As Int32 = Int32.MinValue
        '''<summary>Height of every image in pixel.</summary>
        Public FrameHeight As Int32 = Int32.MinValue
        '''<summary>True bit depth per pixel per plane, see http://www.grischa-hahn.homepage.t-online.de/astro/ser/SER%20Doc%20V3b.pdf for details.</summary>
        Public PixelDepthPerPlane As Int32 = Int32.MinValue
        '''<summary>Number of image frames in SER file.</summary>
        Public FrameCount As Int32 = Int32.MinValue
        '''<summary>Name of observer.</summary>
        Public Observer As String = String.Empty
        '''<summary>Name of used camera.</summary>
        Public Instrument As String = String.Empty
        '''<summary>Name of used telescope.</summary>
        Public Telescope As String = String.Empty
        '''<summary>Start time of image stream (local time) - If 12_DateTime <= 0 Then 12_DateTime Is invalid And the SER file does Not contain a time stamp trailer.</summary>
        Public DateTimeLocalRaw As Int64 = Int64.MinValue
        '''<summary>Start time of image stream in UTC.</summary>
        Public DateTimeUTCRaw As Int64 = Int64.MinValue

        Private MyTrailerLength As Int64 = Int64.MinValue

        Public Sub New()

        End Sub

        Public Sub New(ByRef BinaryIN As System.IO.BinaryReader)
            Header = System.Text.Encoding.ASCII.GetString(BinaryIN.ReadBytes(14))
            LuID = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            ColorID = CType(BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0), eColorID)
            LittleEndian = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            FrameWidth = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            FrameHeight = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            PixelDepthPerPlane = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            FrameCount = BitConverter.ToInt32(BinaryIN.ReadBytes(4), 0)
            Observer = System.Text.Encoding.ASCII.GetString(BinaryIN.ReadBytes(40))
            Instrument = System.Text.Encoding.ASCII.GetString(BinaryIN.ReadBytes(40))
            Telescope = System.Text.Encoding.ASCII.GetString(BinaryIN.ReadBytes(40))
            DateTimeLocalRaw = BitConverter.ToInt64(BinaryIN.ReadBytes(8), 0)
            DateTimeUTCRaw = BitConverter.ToInt64(BinaryIN.ReadBytes(8), 0)
            MyTrailerLength = BinaryIN.BaseStream.Length - SERHeaderLength - TotalImageBytes
        End Sub

        Public ReadOnly Property DateTimeSubSec() As UInt64
            Get
                Return CULng(DateTimeLocalRaw - ((DateTimeLocalRaw \ 10000000) * 10000000))
            End Get
        End Property

        Public ReadOnly Property DateTimeLocal() As DateTime
            Get
                Return New DateTime(1, 1, 1, 0, 0, 0).AddSeconds(DateTimeLocalRaw \ 10000000)
            End Get
        End Property

        Public ReadOnly Property DateTimeUTC() As DateTime
            Get
                Return New DateTime(1, 1, 1, 0, 0, 0).AddSeconds(DateTimeUTCRaw \ 10000000)
            End Get
        End Property

        Public ReadOnly Property BytePerPixel() As Integer
            Get
                Return PixelDepthPerPlane \ 8
            End Get
        End Property

        Public ReadOnly Property TotalImageBytes() As Long
            Get
                Return CLng(FrameCount) * CLng(FrameWidth) * CLng(FrameHeight) * CLng(BytePerPixel)
            End Get
        End Property

        Public ReadOnly Property TrailerLength() As Long
            Get
                Return MyTrailerLength
            End Get
        End Property

        Public Function PrintInfo() As List(Of String)
            Dim RetVal As New List(Of String)
            RetVal.Add("Header             " & Header)
            RetVal.Add("LuID               " & LuID.ValRegIndep)
            RetVal.Add("ColorID            " & ColorID.ToString.Trim)
            RetVal.Add("LittleEndian       " & LittleEndian.ValRegIndep)
            RetVal.Add("ImageWidth         " & FrameWidth.ValRegIndep & " pixel")
            RetVal.Add("ImageHeight        " & FrameHeight.ValRegIndep & " pixel")
            RetVal.Add("PixelDepthPerPlane " & PixelDepthPerPlane.ValRegIndep)
            RetVal.Add("FrameCount         " & FrameCount.ValRegIndep)
            RetVal.Add("Observer           " & Observer)
            RetVal.Add("Instrument         " & Instrument)
            RetVal.Add("Telescope          " & Telescope)
            RetVal.Add("DateTimeLocalRaw   " & DateTimeLocalRaw.ValRegIndep)
            RetVal.Add("DateTimeUTCRaw     " & DateTimeUTCRaw.ValRegIndep)
            RetVal.Add("DERIVED:")
            RetVal.Add("  DateTimeSubSec   " & DateTimeSubSec.ValRegIndep)
            RetVal.Add("  DateTimeLocal    " & DateTimeLocal.ToString)
            RetVal.Add("  DateTimeUTC      " & DateTimeUTC.ToString)
            RetVal.Add("  BytePerPixel     " & BytePerPixel.ValRegIndep)
            RetVal.Add("  TotalImageBytes  " & TotalImageBytes.ValRegIndep)
            RetVal.Add("  TrailerLength    " & TrailerLength.ValRegIndep)
            Return RetVal
        End Function

    End Class

    Public Class cSerFormatWriter

        Dim FileIO As System.IO.FileStream
        Dim BinaryOUT As System.IO.BinaryWriter

        Public Property Header As New cSERHeader

        Public Sub InitForWrite(ByVal NewSERFile As String)

            FileIO = New System.IO.FileStream(NewSERFile, IO.FileMode.Create, IO.FileAccess.Write)
            BinaryOUT = New System.IO.BinaryWriter(FileIO)

            BinaryOUT.Write(System.Text.Encoding.ASCII.GetBytes("LUCAM-RECORDER"))                  'Header
            BinaryOUT.Write(BitConverter.GetBytes(CType(4660, Int32)))                              'LuID
            BinaryOUT.Write(BitConverter.GetBytes(CType(0, Int32)))                                 'ColorID
            BinaryOUT.Write(BitConverter.GetBytes(CType(0, Int32)))                                 'LittleEndian
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.FrameWidth, Int32)))                 'ImageWidth
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.FrameHeight, Int32)))                'ImageHeight
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.PixelDepthPerPlane, Int32)))         'PixelDepthPerPlane
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.FrameCount, Int32)))                 'FrameCount
            BinaryOUT.Write(System.Text.Encoding.ASCII.GetBytes(Header.Observer.PadRight(40)))      'Observer
            BinaryOUT.Write(System.Text.Encoding.ASCII.GetBytes(Header.Instrument.PadRight(40)))    'Instrument
            BinaryOUT.Write(System.Text.Encoding.ASCII.GetBytes(Header.Telescope.PadRight(40)))     'Telescope
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.DateTimeLocalRaw, Int64)))           'FrameCount
            BinaryOUT.Write(BitConverter.GetBytes(CType(Header.DateTimeUTCRaw, Int64)))             'FrameCount

        End Sub

        Public Sub AppendFrame(ByRef Frame(,) As UInt16)
            For Idx1 As Integer = 0 To Frame.GetUpperBound(0)
                For Idx2 As Integer = 0 To Frame.GetUpperBound(1)
                    BinaryOUT.Write(BitConverter.GetBytes(Frame(Idx1, Idx2)))
                Next Idx2
            Next Idx1
        End Sub

        Public Sub CloseSerFile()
            BinaryOUT.Flush()
            FileIO.Flush()
            BinaryOUT.Close()
            FileIO.Close()
        End Sub

    End Class


End Class