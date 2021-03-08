Option Explicit On
Option Strict On

'''<summary>Format helper and converter for the SER format handling.</summary>
'''<see cref="https://free-astro.org/index.php/SER"/>
Partial Public Class cSERFormat

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
        '''<summary>SER time has a resolution of 100 ns, so this divider generates seconds from SER times.</summary>
        Public Const SERTimeToSecDivider As UInt64 = 10000000
        '''<summary>Moment 0 in the SER timestamp meaning.</summary>
        Public ReadOnly Property SERTimeZero As New DateTime(1, 1, 1, 0, 0, 0)

        '''<summary>"LUCAM-RECORDER".</summary>
        Public Header As String = String.Empty
        '''<summary>Lumenera camera series ID - unused, default is 0.</summary>
        Public LuID As Int32 = Int32.MinValue
        '''<summary>Color ID.</summary>
        Public ColorID As eColorID
        '''<summary>0 for big-endian, 1 for little-endian byte order in 16 bit image data.</summary>
        Public LittleEndian As Int32 = Int32.MinValue
        '''<summary>Width of every image in pixel.</summary>
        Public FrameWidth As Int32 = 0
        '''<summary>Height of every image in pixel.</summary>
        Public FrameHeight As Int32 = 0
        '''<summary>True bit depth per pixel per plane, see http://www.grischa-hahn.homepage.t-online.de/astro/ser/SER%20Doc%20V3b.pdf for details.</summary>
        Public PixelDepthPerPlane As Int32 = 0
        '''<summary>Number of image frames in SER file.</summary>
        Public FrameCount As Int32 = 0
        '''<summary>Name of observer.</summary>
        Public Observer As String = String.Empty
        '''<summary>Name of used camera.</summary>
        Public Instrument As String = String.Empty
        '''<summary>Name of used telescope.</summary>
        Public Telescope As String = String.Empty
        '''<summary>Start time of image stream (local time) - If 12_DateTime <= 0 Then 12_DateTime Is invalid And the SER file does Not contain a time stamp trailer.</summary>
        Public DateTimeLocalRaw As Int64 = 0
        '''<summary>Start time of image stream in UTC.</summary>
        Public DateTimeUTCRaw As Int64 = 0

        '''<summary>Raw data of the trailer.</summary>
        Public TrailerSeconds() As UInt64 = {}

        Public Sub New()
            MyValidSERFile = False
        End Sub

        '''<summary>Start reading the passed binary reader stream as a SER file.</summary>
        '''<param name="BinaryIN">Binary stream.</param>
        Public Sub New(ByRef BinaryIN As System.IO.BinaryReader)
            'Check if the size is at least the header size
            MyValidSERFile = False
            If BinaryIN.BaseStream.Length < SERHeaderLength Then Exit Sub
            BinaryIN.BaseStream.Seek(0, IO.SeekOrigin.Begin)
            'Read the header
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
            'Validate if the file is valid
            Dim ExpectedTrailerSize As Long = FrameCount * 8
            Dim StreamLength As Long = BinaryIN.BaseStream.Length
            If StreamLength = SERHeaderLength + TotalImageBytes + ExpectedTrailerSize Then
                'Full SER file with time
                MyFileFrames = FrameCount
                MyTrailerPresent = True
                MyValidSERFile = True
            Else
                If StreamLength = SERHeaderLength + TotalImageBytes Then
                    'Full SER file without time
                    MyFileFrames = FrameCount
                    MyTrailerPresent = False
                    MyValidSERFile = True
                Else
                    'Cut SER file
                    MyFileFrames = CInt(Math.Floor((StreamLength - SERHeaderLength) / BytePerFrame))
                    MyTrailerPresent = False
                    MyValidSERFile = True
                End If
            End If
        End Sub

        '''<summary>Determin if the loaded SER file is a valid SER file.</summary>
        Public ReadOnly Property ValidSERFile As Boolean = MyValidSERFile
        Private MyValidSERFile As Boolean = False

        '''<summary>Number of frames according to the file size (may not be the same as the header indicates).</summary>
        Public ReadOnly Property FileFrames As Integer
            Get
                Return MyFileFrames
            End Get
        End Property
        Private MyFileFrames As Integer = 0

        '''<summary>Is a valid and full trailer present?.</summary>
        Public ReadOnly Property TrailerPresent As Boolean
            Get
                Return MyTrailerPresent
            End Get
        End Property
        Private MyTrailerPresent As Boolean = False

        Public ReadOnly Property DateTimeSubSec As UInt64
            Get
                Return CULng(DateTimeLocalRaw - ((DateTimeLocalRaw / SERTimeToSecDivider) * SERTimeToSecDivider))
            End Get
        End Property

        Public ReadOnly Property DateTimeLocal As DateTime
            Get
                Return SerTimestampToDateTime(DateTimeLocalRaw)
            End Get
        End Property

        '''<summary>Start time of image stream in UTC.</summary>
        Public ReadOnly Property DateTimeUTC As DateTime
            Get
                Return SerTimestampToDateTime(DateTimeUTCRaw)
            End Get
        End Property

        '''<summary>Bytes per pixel.</summary>
        Public ReadOnly Property BytePerPixel As Integer
            Get
                Return PixelDepthPerPlane \ 8
            End Get
        End Property

        '''<summary>Number of bytes a full frame consumes.</summary>
        Public ReadOnly Property BytePerFrame As Long
            Get
                Return CLng(FrameWidth) * CLng(FrameHeight) * CLng(BytePerPixel)
            End Get
        End Property

        '''<summary>Number of bytes all frame consumes (according to the header information).</summary>
        Public ReadOnly Property TotalImageBytes As Long
            Get
                Return CLng(FrameCount) * BytePerFrame
            End Get
        End Property

        '''<summary>Length of the trailer [byte].</summary>
        Public ReadOnly Property TrailerLength As Long
            Get
                Return MyTrailerLength
            End Get
        End Property
        Private MyTrailerLength As Int64 = 0

        ''<summary>Convert a SER formated time stamp to a "common" DateTime value.</summary>
        Public Function SerTimestampToDateTime(ByVal TimeStamp As Int64) As DateTime
            Return SERTimeZero.AddSeconds(TimeStamp / SERTimeToSecDivider)
        End Function

        ''<summary>Get a list of all time stamps.</summary>
        Public Sub ReadTrailer(ByRef BinaryIN As System.IO.BinaryReader)
            Dim OldPos As Long = BinaryIN.BaseStream.Position
            BinaryIN.BaseStream.Seek(BinaryIN.BaseStream.Length - TrailerLength, IO.SeekOrigin.Begin)
            Dim TrailerSeconds(FrameCount - 1) As UInt64
            For Ptr As Integer = 0 To FrameCount - 1
                TrailerSeconds(Ptr) = BitConverter.ToUInt64(BinaryIN.ReadBytes(8), 0)
                Dim Seconds As Double = RawDate / SERTimeToSecDivider
                RetVal(Ptr) = SERTimeZero.AddSeconds(Seconds)
            Next Ptr
            'Restore old stream position and return list of DateTime
            BinaryIN.BaseStream.Seek(OldPos, IO.SeekOrigin.Begin)
        End Sub

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

End Class