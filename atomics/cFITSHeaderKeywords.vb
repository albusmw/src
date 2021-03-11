Option Explicit On
Option Strict On

'''<summary>Specific FITS parameters.</summary>
Public Class FITSSpec
    '''<summary>Length of one header element.</summary>
    Public Const HeaderElementLength As Integer = 80
    '''<summary>Length of a header block - FITS files may contain an integer size of header blocks.</summary>
    Public Const HeaderBlockSize As Integer = 2880
    '''<summary>Number of header elements per header block.</summary>
    Public Const HeaderElements As Integer = 36  'HeaderBlockSize \ HeaderElementLength
    '''<summary>Length of a keyword (without "=").</summary>
    Public Const HeaderKeywordLength As Integer = 8
    '''<summary>Keyword-value separator - always equal sign followed by a space.</summary>
    Public Const HeaderEqualString As String = "= "
    '''<summary>Typical value length, taken from a MAXIM file; a space and a \ sign may follow for a comment.</summary>
    Public Const HeaderValueLength As Integer = 20
    '''<summary>ASCII code for a space sign.</summary>
    Public Const HeaderSpaceCode As Byte = 32
    '''<summary>Ensure that the given card has the correct length.</summary>
    Public Shared Function EnsureCorrectLength(ByVal NewCard As String) As String
        If NewCard.Length > FITSSpec.HeaderElementLength Then
            Return NewCard.Substring(0, FITSSpec.HeaderElementLength)
        Else
            Return NewCard.PadRight(FITSSpec.HeaderElementLength)
        End If
    End Function
    '''<summary>Ensure that the given card has the correct length.</summary>
    Public Shared Function EnsureCorrectLength(ByVal NewCard As List(Of Byte)) As List(Of Byte)
        If NewCard.Count > FITSSpec.HeaderElementLength Then
            Return NewCard.GetRange(0, FITSSpec.HeaderElementLength)
        Else
            Return NewCard.Concat(Enumerable.Repeat(Of Byte)(HeaderSpaceCode, FITSSpec.HeaderElementLength - NewCard.Count)).ToList
        End If
    End Function
End Class

'================================================================================
'''<summary>Attribute used for hold the FITS keyword.</summary>
<AttributeUsage(AttributeTargets.All, Inherited:=True, AllowMultiple:=False)>
Public Class FITSKeyword
    Inherits System.Attribute
    Public Sub New()
        Me.New(String.Empty)
    End Sub
    Public Sub New(ByVal Keyword As String)
        MyKeywords = New List(Of String)({Keyword})
    End Sub
    Public Sub New(ByVal Keywords As String())
        MyKeywords = New List(Of String)
        MyKeywords.AddRange(Keywords)
    End Sub
    Public ReadOnly Property Keywords() As List(Of String)
        Get
            Return MyKeywords
        End Get
    End Property
    Private MyKeywords As List(Of String)
    '''<summary>The keyword string(s) associated with the given keyword.</summary>
    Public Shared Function GetKeywords(ByRef Element As eFITSKeywords) As String()
        Dim attributes As FITSKeyword() = CType(Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(FITSKeyword), False), FITSKeyword())
        If attributes.Length > 0 Then Return attributes(0).Keywords.ToArray Else Return New String() {}
    End Function
    '''<summary>The (common) keyword string associated with the given keyword.</summary>
    Public Shared Function KeywordString(ByRef Element As eFITSKeywords) As String
        Dim attributes As FITSKeyword() = CType(Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(FITSKeyword), False), FITSKeyword())
        If attributes.Length > 0 Then Return attributes(0).Keywords(0) Else Return String.Empty
    End Function
End Class
'================================================================================

'================================================================================
'''<summary>Attribute used for hold the FITS keyword.</summary>
<AttributeUsage(AttributeTargets.All, Inherited:=True, AllowMultiple:=False)>
Public Class FITSComment
    Inherits System.Attribute
    Public Sub New()
        Me.New(String.Empty)
    End Sub
    Public Sub New(ByVal Comment As String)
        MyComment = Comment
    End Sub
    Public ReadOnly Property Comment() As String
        Get
            Return MyComment
        End Get
    End Property
    Private MyComment As String = String.Empty
    '''<summary>The comment string associated with the given keyword.</summary>
    Public Shared Function GetComment(ByRef Element As eFITSKeywords) As String
        Dim attributes As FITSComment() = CType(Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(FITSComment), False), FITSComment())
        If attributes.Length > 0 Then Return attributes(0).Comment Else Return String.Empty
    End Function
End Class

'================================================================================

Public Class cFITSType

    '<summary>Describes which data type a certain keyword has.</summary>
    'Public Shared Function GetDataType(ByVal Keyword As eFITSKeywords) As String
    '    'Set the correct data type
    '    Select Case Keyword
    '        Case eFITSKeywords.BITPIX, eFITSKeywords.NAXIS, eFITSKeywords.NAXIS1, eFITSKeywords.NAXIS2
    '            Return "INTEGER"
    '        Case eFITSKeywords.BZERO, eFITSKeywords.BSCALE, eFITSKeywords.FOV1, eFITSKeywords.FOV2, eFITSKeywords.PIXSIZE1, eFITSKeywords.PIXSIZE2, eFITSKeywords.PLATESZ1, eFITSKeywords.PLATESZ2
    '            Return "DOUBLE"
    '        Case Else
    '            Return String.Empty
    '    End Select
    'End Function

    '''<summary>Get the FITS compatible formated string.</summary>
    Public Shared Function AsString(ByVal Value As Object) As String
        If TypeOf Value Is String Then Return FITSString(CStr(Value))
        If TypeOf Value Is Single Then Return FITSString(CDbl(Value))
        If TypeOf Value Is Double Then Return FITSString(CDbl(Value))
        If TypeOf Value Is Byte Then Return CStr(Value).Trim
        If TypeOf Value Is Int16 Then Return CStr(Value).Trim
        If TypeOf Value Is Int32 Then Return CStr(Value).Trim
        If TypeOf Value Is Int64 Then Return CStr(Value).Trim
        If TypeOf Value Is UInt16 Then Return CStr(Value).Trim
        If TypeOf Value Is UInt32 Then Return CStr(Value).Trim
        If TypeOf Value Is UInt64 Then Return CStr(Value).Trim
        If TypeOf Value Is DateTime Then Return FITSString(CType(Value, DateTime))
        If TypeOf Value Is Date Then Return FITSString(CType(Value, Date))
        If TypeOf Value Is TimeSpan Then Return FITSString(CType(Value, TimeSpan))
        Return CStr(Value)
    End Function

    '''<summary>Formated content as a string.</summary>
    Public Shared Function FITSString(ByVal Value As String) As String
        Return "'" & Value & "'"
    End Function

    '''<summary>Formated content as a double number.</summary>
    Public Shared Function FITSString(ByVal Value As Double) As String
        Return Value.ToString.Trim.Replace(",", ".")
    End Function

    '''<summary>Formated content for all "DATE..." fields, time.</summary>
    Public Shared Function FITSString(ByVal Moment As DateTime) As String
        If Moment.Hour = 0 And Moment.Minute = 0 And Moment.Second = 0 And Moment.Millisecond = 0 Then
            Return FITSString(Format(Moment, "yyyy-MM-dd"))
        Else
            Return FITSString(Format(Moment, "yyyy-dd-MMTHH:mm:ss"))
        End If
    End Function

    '''<summary>Formated content for all "TIME..." fields, time.</summary>
    Public Shared Function FITSString(ByVal Moment As TimeSpan) As String
        Return FITSString(Format(Moment.Hours, "00") & ":" & Format(Moment.Minutes, "00") & ":" & Format(Moment.Seconds, "00") & "." & Format(Moment.Milliseconds, "000"))
    End Function

    '''<summary>Formated content for all "TIME..." fields, time.</summary>
    Public Shared Function FITSString_DateTime(ByVal Moment As DateTime) As String
        Return FITSString(Format(Moment, "yyyy-MM-dd HH:mm:ss.fff"))
    End Function

End Class

'================================================================================

'''<summary>Class to provide FITS header keywords (elements and service functions).</summary>
'''<see cref="http://wise2.ipac.caltech.edu/docs/release/prelim/expsup/sec2_3b.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/standard_dict.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
'''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
'''<see cref="http://eso-python.github.io/ESOPythonTutorials/FITS-images.html"/>
'''<see cref="https://lco.global/documentation/data/fits-headers/"/>
'''<see cref="https://proba2.sidc.be/data/SWAP/level0"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/heasarc/fits/java/v1.0/javadoc/nom/tam/fits/header/extra/SBFitsExt.html"/>
Public Enum eFITSKeywords

    '''<summary>Enum value is unknown.</summary>
    <FITSKeyword("[UNKNOWN]")>
    <FITSComment("")>
    [UNKNOWN]

    '''<summary>Altitude axis position.</summary>
    <FITSKeyword("ALTITUDE")>
    <FITSComment("Altitude axis position.")>
    [ALTITUDE]

    '''<summary>The value field shall contain a character string identifying who compiled the information In the data associated With the header. This keyword Is appropriate When the data originate In a published paper Or are compiled from many sources.</summary>
    <FITSKeyword("AUTHOR")>
    <FITSComment("")>
    [AUTHOR]

    '''<summary>Altitude axis position.</summary>
    <FITSKeyword("AZIMUTH")>
    <FITSComment("Azimuth axis position.")>
    [AZIMUTH]

    '''<summary>.</summary>
    <FITSKeyword("BAYOFFX")>
    <FITSComment("")>
    [BAYOFFX]

    '''<summary>.</summary>
    <FITSKeyword("BAYOFFY")>
    <FITSComment("")>
    [BAYOFFY]

    '''<summary>If present the image has a valid Bayer color pattern.</summary>
    <FITSKeyword("BAYERPAT")>
    <FITSComment("")>
    [BAYERPAT]

    '''<summary>8 unsigned int, 16 and 32 int, -32 and -64 real.</summary>
    <FITSKeyword("BITPIX")>
    <FITSComment("8 unsigned int, 16 & 32 int, -32 & -64 real")>
    [BITPIX]

    '''<summary></summary>
    <FITSKeyword("BLKLEVEL")>
    <FITSComment("??????")>
    [BLKLEVEL]

    '''<summary>Configured BRIGHTNESS value of the camera.</summary>
    <FITSKeyword("BRIGHTN")>
    <FITSComment("Configured BRIGHTNESS value of the camera")>
    [BRIGHTNESS]

    '''<summary>Scaling factor in scaling equation.</summary>
    <FITSKeyword("BSCALE")>
    <FITSComment("physical = BZERO + BSCALE*array_value")>
    [BSCALE]

    '''<summary>Zero point in scaling equation.</summary>
    <FITSKeyword("BZERO")>
    <FITSComment("physical = BZERO + BSCALE*array_value")>
    [BZERO]

    '''<summary>Actual measured sensor temperature at the start of exposure in degrees C. Absent if temperature is not available.</summary>
    <FITSKeyword("CCD-TEMP")>
    <FITSComment("CCD temperature at start of exposure in C")>
    [CCDTEMP]

    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 1 pixel scale at CRPIX1,CRPIX2. See PXSCAL1 For arcsec equivalent.</remarks>
    '''<example>-0.0003819444391411 degrees/pix</example>
    <FITSKeyword("CDELT1")>
    <FITSComment("Axis 1 pixel scale at CRPIX1,CRPIX2")>
    [CDELT1]

    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 2 pixel scale at CRPIX1,CRPIX2. See PXSCAL2 For arcsec equivalent.</remarks>
    '''<example>0.0003819444391411 degrees/pix</example>
    <FITSKeyword("CDELT2")>
    <FITSComment("Axis 2 pixel scale at CRPIX1,CRPIX2")>
    [CDELT2]

    '''<summary>Nominal Altitude of center of image in degrees.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("CENTALT")>
    <FITSComment("Nominal Altitude of center of image in degrees")>
    [CENTALT]

    '''<summary>Nominal Azimuth of center of image in degrees.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("CENTAZ")>
    <FITSComment("Nominal Azimuth of center of image in degrees")>
    [CENTAZ]

    '''<summary>Type of color sensor Bayer array or zero for monochrome.</summary>
    <FITSKeyword("COLORTYP")>
    <FITSComment("")>
    [COLORTYP]

    '''<summary>.</summary>
    <FITSKeyword("COMMENT")>
    <FITSComment("")>
    [COMMENT]

    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    <FITSKeyword("CROTA1")>
    <FITSComment("Rotation [degree] from a standard coordinate")>
    [CROTA1]

    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    <FITSKeyword("CROTA2")>
    <FITSComment("Rotation [degree] from a standard coordinate")>
    [CROTA2]

    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS1+1)</remarks>
    '''<example>2048.000000</example>
    <FITSKeyword("CRPIX1")>
    <FITSComment("Axis 1 reference pixel at CRVAL1,CRVAL2")>
    [CRPIX1]

    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS2+1)</remarks>
    '''<example>2048.000000</example>
    <FITSKeyword("CRPIX2")>
    <FITSComment("Axis 2 reference pixel at CRVAL1,CRVAL2")>
    [CRPIX2]

    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Right Ascension at CRPIX1,CRPIX2 for EQUINOX. </remarks>
    <FITSKeyword("CRVAL1")>
    <FITSComment("Right Ascension at CRPIX1,CRPIX2 for EQUINOX")>
    [CRVAL1]

    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    '''<remarks> Declination at CRPIX1,CRPIX2 for EQUINOX.</remarks>
    <FITSKeyword("CRVAL2")>
    <FITSComment("Declination at CRPIX1,CRPIX2 for EQUINOX")>
    [CRVAL2]

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta and Greisen, 2002</summary>
    '''<example>'RA---SIN'</example>
    <FITSKeyword("CTYPE1")>
    <FITSComment("Projection type for axis 1.")>
    [CTYPE1]

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta and Greisen, 2002</summary>
    '''<example>'DEC--SIN'</example>
    <FITSKeyword("CTYPE2")>
    <FITSComment("Projection type for axis 2.")>
    [CTYPE2]

    '''<summary>Used to color encoding.</summary>
    <FITSKeyword("CTYPE3")>
    <FITSComment("Used to color encoding.")>
    [CTYPE3]

    '''<summary>.</summary>
    <FITSKeyword({"DATE"})>
    <FITSComment("")>
    [DATE]

    '''<summary>The value field shall contain a character string that gives the date on which the observation ended, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword({"DATE_END", "DATE-END"})>
    <FITSComment("Observation end date, UT")>
    [DATE_END]

    '''<summary>The value field shall contain a character string that gives the date on which the observation started, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword({"DATE_OBS", "DATE-OBS"})>
    <FITSComment("Observation start date, UT")>
    [DATE_OBS]

    '''<summary>The value field gives the declination of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'dd:mm:ss.sss' format where the decimal point and number of fractional digits are optional.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("DEC")>
    <FITSComment("Declination of the observed object")>
    [DEC]

    '''<summary>The value field shall contain a floating point number giving the nominal declination of the pointing direction in units of decimal degrees. The coordinate reference frame is given by the RADECSYS keyword, and the coordinate epoch is given by the EQUINOX keyword. The precise definition of this keyword is instrument-specific, but typically the nominal direction corresponds to the direction to which the instrument was requested to point. The DEC_PNT keyword should be used to give the actual pointed direction.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("DEC_NOM")>
    <FITSComment("Nominal declination of the observation")>
    [DEC_NOM]

    '''<summary>The value field shall contain a floating point number giving the declination of the observed object in units of decimal degrees. The coordinate reference frame is given by the RADECSYS keyword, and the coordinate epoch is given by the EQUINOX keyword.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("DEC_OBJ")>
    <FITSComment("Declination of the observed object")>
    [DEC_OBJ]

    '''<summary>Electronic gain in photoelectrons per ADU.</summary>
    <FITSKeyword("EGAIN")>
    <FITSComment("Electronic gain - photoelectrons per ADU")>
    [EGAIN]

    '''<summary>Primary HDU.</summary>
    <FITSKeyword("END")>
    <FITSComment("")>
    [END]

    '''<summary>The value field shall contain a floating point number giving the exposure time of the observation in units of seconds.</summary>
    <FITSKeyword("EXPTIME")>
    <FITSComment("Exposure time in seconds")>
    [EXPTIME]

    '''<summary></summary>
    <FITSKeyword("EXTEND")>
    <FITSComment("")>
    [EXTEND]

    '''<summary>Equinox of the World Coordinate System (WCS).</summary>
    '''<example>2000.0</example>
    <FITSKeyword("EQUINOX")>
    <FITSComment("Equinox of the World Coordinate System (WCS)")>
    [EQUINOX]

    '''<summary>Name of selected filter.</summary>
    <FITSKeyword("FILTER")>
    <FITSComment("Name of selected filter")>
    [FILTER]

    '''<summary>Focus value (from logbook). Used when a single value is given in the logs.</summary>
    <FITSKeyword("FOCUS")>
    <FITSComment("Focuser position in steps")>
    [FOCUS]

    '''<summary>Focuser temperature readout in degrees C, if available.</summary>
    <FITSKeyword("FOCUSTEM")>
    <FITSComment("Focuser temperature readout in degrees C")>
    [FOCUSTEM]

    '''<summary>Field of view [°] along axis 1.</summary>
    <FITSKeyword("FOV1")>
    <FITSComment("Field of view [degree] along axis 1.")>
    [FOV1]

    '''<summary>Field of view [°] along axis 2.</summary>
    <FITSKeyword("FOV2")>
    <FITSComment("Field of view [degree] along axis 2.")>
    [FOV2]

    '''<summary>Configured GAIN value of the camera.</summary>
    <FITSKeyword("GAIN")>
    <FITSComment("GAIN value of the camera")>
    [GAIN]

    '''<summary>Type of image: Light Frame, Bias Frame, Dark Frame, Flat Frame, or Tricolor Image.</summary>
    <FITSKeyword("IMAGETYP")>
    <FITSComment("Type of image")>
    [IMAGETYP]

    '''<summary>The value field shall contain a character string identifying the instrument used to acquire the data associated with the header.</summary>
    <FITSKeyword("INSTRUME")>
    <FITSComment("")>
    [INSTRUME]

    '''<summary>Primary HDU - Number of data axes. Always = 2 for two-dimensional images.</summary>
    <FITSKeyword("NAXIS")>
    <FITSComment("Number of axes")>
    [NAXIS]

    '''<summary>Primary HDU - Length of data axis 1 or number of columns in image.</summary>
    <FITSKeyword("NAXIS1")>
    <FITSComment("Fastest changing axis")>
    [NAXIS1]

    '''<summary>Primary HDU - Length of data axis 2 or number of rows in image.</summary>
    <FITSKeyword("NAXIS2")>
    <FITSComment("Next to fastest changing axis")>
    [NAXIS2]

    '''<summary>Primary HDU - Length of data axis 3 or number of color channels in image.</summary>
    <FITSKeyword("NAXIS3")>
    <FITSComment("Length of axis 3 or number of color channels")>
    [NAXIS3]

    '''<summary>This is the Right Ascension of the center of the image in hours, minutes and secon ds. The format for this is 12 24 23.123 (HH MM SS.SSS) using a space as the separator.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/heasarc/fits/java/v1.0/javadoc/nom/tam/fits/header/extra/SBFitsExt.html"/>
    <FITSKeyword("OBJCTRA")>
    <FITSComment("This is the Right Ascension of the center of the image in hours.")>
    [OBJCTRA]

    '''<summary>This is the Declination of the center of the image in degrees. The format for this is +25 12 34.111 (SDD MM SS.SSS) using a space as the separator. For the sign, North is + and South is -.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/heasarc/fits/java/v1.0/javadoc/nom/tam/fits/header/extra/SBFitsExt.html"/>
    <FITSKeyword("OBJCTDEC")>
    <FITSComment("This is the Declination of the center of the image in degrees.")>
    [OBJCTDEC]

    '''<summary>Nominal Altitude of center of image.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("OBJCTALT")>
    <FITSComment("Nominal Altitude of center of image")>
    [OBJCTALT]

    '''<summary>Nominal Azimuth of center of image.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("OBJCTAZ")>
    <FITSComment("Nominal Azimuth of center of image")>
    [OBJCTAZ]

    '''<summary>The value field shall contain a character string giving a name for the object observed.</summary>
    <FITSKeyword("OBJECT")>
    <FITSComment("")>
    [OBJECT]

    '''<summary>The value field shall contain a character string giving a name for the observed object that conforms to the IAU astronomical Object naming conventions. The value of this keyword Is more strictly constrained than For the standard Object keyword which In practice has often been used To record other ancillary information about the observation (e.g. filter, exposure time, weather conditions, etc.).</summary>
    <FITSKeyword("OBJNAME")>
    <FITSComment("")>
    [OBJNAME]

    '''<summary>The value field shall contain a character string identifying who acquired the data associated with the header.</summary>
    <FITSKeyword("OBSERVER")>
    <FITSComment("")>
    [OBSERVER]

    ''''<summary>The value field shall contain a character string which uniquely identifies the dataset contained In the FITS file. This Is typically a sequence number that can contain a mixture Of numerical And character values. Example '10315-01-01-30A'.</summary>
    <FITSKeyword({"OBS_ID", "OBS-ID"})>
    <FITSComment("")>
    [OBS_ID]

    '''<summary>Configured OFFSET value of the camera.</summary>
    <FITSKeyword("OFFSET")>
    <FITSComment("Configured OFFSET value of the camera")>
    [OFFSET]

    '''<summary>The value field shall contain a character string identifying the organization or institution responsible for creating the FITS file.</summary>
    <FITSKeyword("ORIGIN")>
    <FITSComment("organization or institution created the FITS file")>
    [ORIGIN]

    '''<summary>Pixel size [um] along axis 1.</summary>
    <FITSKeyword("PIXSIZE1")>
    <FITSComment("Pixel size along axis 1 after binning - um")>
    [PIXSIZE1]

    '''<summary>Pixel size [um] along axis 2.</summary>
    <FITSKeyword("PIXSIZE2")>
    <FITSComment("Pixel size along axis 2 after binning - um")>
    [PIXSIZE2]

    ''<summary>Plate size [cm] along axis 1.</summary>
    <FITSKeyword("PLATESZ1")>
    <FITSComment("Plate size [cm] along axis 1")>
    [PLATESZ1]

    ''<summary>Plate size [cm] along axis 2.</summary>
    <FITSKeyword("PLATESZ2")>
    <FITSComment("Plate size [cm] along axis 2")>
    [PLATESZ2]

    '''<summary>The value field shall contain a character string giving the name, And optionally, the version of the program that originally created the current FITS HDU. This keyword Is synonymous With the CREATOR keyword.  Example 'TASKNAME V1.2.3'.</summary>
    <FITSKeyword("PROGRAM")>
    <FITSComment("Program that originally created the current FITS HDU")>
    [PROGRAM]

    '''<summary>Pixel scale at CRPIX1,CRPIX2 for axis1.</summary>
    '''<example>1.375 arcsec/pix</example>
    <FITSKeyword("PXSCAL1")>
    <FITSComment("Pixel scale at CRPIX1,CRPIX2 for axis1")>
    [PXSCAL1]

    '''<summary>Pixel scale at CRPIX1,CRPIX2 for axis2.</summary>
    '''<example>1.375 arcsec/pix</example>
    <FITSKeyword("PXSCAL21")>
    <FITSComment("Pixel scale at CRPIX1,CRPIX2 for axis2")>
    [PXSCAL2]

    '''<summary>QHY read-out mode.</summary>
    <FITSKeyword("QHY_MODE")>
    <FITSComment("QHY read-out mode")>
    [QHY_MODE]

    '''<summary>The value field gives the Right Ascension of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'HH:MM:SS.sss' format where the decimal point and number of fractional digits are optional. The coordinate reference frame is given by the RADECSYS keyword, and the coordinate epoch is given by the EQUINOX keyword. Example: 180.6904 or '12:02:45.7'.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("RA")>
    <FITSComment("Right ascension of the observed object")>
    [RA]

    '''<summary>The value field shall contain a floating point number giving the nominal Right Ascension of the pointing direction in units of decimal degrees. The coordinate reference frame is given by the RADECSYS keyword, and the coordinate epoch is given by the EQUINOX keyword. The precise definition of this keyword is instrument-specific, but typically the nominal direction corresponds to the direction to which the instrument was requested to point. The RA_PNT keyword should be used to give the actual pointed direction.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("DEC_NOM")>
    <FITSComment("Nominal right ascension of the observation")>
    [RA_NOM]

    '''<summary>The value field shall contain a floating point number giving the Right Ascension of the observed object in units of decimal degrees. The coordinate reference frame is given by the RADECSYS keyword, and the coordinate epoch is given by the EQUINOX keyword.</summary>
    '''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
    <FITSKeyword("RA_OBJ")>
    <FITSComment("Right ascension of the observed object")>
    [RA_OBJ]

    '''<summary>CCD temperature setpoint in degrees C. Absent if setpoint was not entered.</summary>
    <FITSKeyword("SET-TEMP")>
    <FITSComment("CCD temperature setpoint in C")>
    [SETTEMP]

    '''<summary>Primary HDU.</summary>
    <FITSKeyword("SIMPLE")>
    <FITSComment("")>
    [SIMPLE]

    '''<summary>The value field shall contain a floating point number giving the geographic latitude from which the observation was made in units of degrees.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("SITELAT")>
    <FITSComment("Geographic latitude from which the observation was made - degree")>
    [SITELAT]

    '''<summary>The value field shall contain a floating point number giving the geographic longitude from which the observation was made in units of degrees.</summary>
    '''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
    <FITSKeyword("SITELONG")>
    <FITSComment("Geographic longitude from which the observation was made - degree")>
    [SITELONG]

    '''<summary></summary>
    <FITSKeyword("SWCREATE")>
    <FITSComment("")>
    [SWCREATE]

    '''<summary>Clear aperture of the telescope [m].</summary>
    <FITSKeyword("TELAPER")>
    <FITSComment("Clear aperture of the telescope -m")>
    [TELAPER]

    '''<summary>The value field shall contain a character string identifying the telescope used to acquire the data associated with the header.</summary>
    <FITSKeyword("TELESCOP")>
    <FITSComment("Telescope used")>
    [TELESCOP]

    ''<summary>Focal length of the telescope [m].</summary>
    <FITSKeyword("TELFOC")>
    <FITSComment("Focal length of the telescope - m")>
    [TELFOC]

    '''<summary>Plate scale of the telescope [arcsec/mm].</summary>
    <FITSKeyword("TELSCALE")>
    <FITSComment("Plate scale of the telescope - arcsec/mm")>
    [TELSCALE]

    '''<summary>The value field shall contain a character string that gives the time at which the observation ended, format 'hh:mm:ss.sss'.</summary>
    <FITSKeyword({"TIME_END", "TIME-END"})>
    <FITSComment("Observation time end, UT")>
    [TIME_END]

    ''''<summary>The value field shall contain a character string that gives the time at which the observation started, format 'hh:mm:ss.sss'.</summary>
    <FITSKeyword({"TIME_OBS", "TIME-OBS"})>
    <FITSComment("Observation time start, UT")>
    [TIME_OBS]

    ''''<summary>Binning factor in width.</summary>
    <FITSKeyword("XBINNING")>
    <FITSComment("Binning factor in width")>
    [XBINNING]

    '''<summary>Pixel size [um] along axis 1.</summary>
    <FITSKeyword("XPIXSZ")>
    <FITSComment("Pixel Width in microns (after binning)")>
    [XPIXSZ]

    ''''<summary>Binning factor in width.</summary>
    <FITSKeyword("YBINNING")>
    <FITSComment("Binning factor in height")>
    [YBINNING]

    '''<summary>Pixel size [um] along axis 2.</summary>
    <FITSKeyword("YPIXSZ")>
    <FITSComment("Pixel Height in microns (after binning)")>
    [YPIXSZ]

    '=========================================================================================================

    <FITSKeyword("EXPOINUS")>
    [EXPOINUS]

    <FITSKeyword("SWOWNER")>
    [SWOWNER]

    <FITSKeyword("DATAMIN")>
    [DATAMIN]

    <FITSKeyword("DATAMAX")>
    [DATAMAX]

    <FITSKeyword("TRUDEPTH")>
    [TRUDEPTH]

    <FITSKeyword("BITSHIFT")>
    [BITSHIFT]

    <FITSKeyword("SUBEXP")>
    [SUBEXP]

    <FITSKeyword("XORGSUBF")>
    [XORGSUBF]

    <FITSKeyword("YORGSUBF")>
    [YORGSUBF]

    <FITSKeyword("FOCALLEN")>
    [FOCALLEN]

    <FITSKeyword("APTDIA")>
    [APTDIA]

    <FITSKeyword("APTAREA")>
    [APTAREA]

    <FITSKeyword("CBLACK")>
    [CBLACK]

    <FITSKeyword("CWHITE")>
    [CWHITE]

    <FITSKeyword("SNAPSHOT")>
    [SNAPSHOT]

    <FITSKeyword("RESMODE")>
    [RESMODE]

    <FITSKeyword("PEDESTAL")>
    [PEDESTAL]

    <FITSKeyword("SBSTDVER")>
    [SBSTDVER]

    <FITSKeyword("SWACQUIR")>
    [SWACQUIR]

    <FITSKeyword("EXPSTATE")>
    [EXPSTATE]

    <FITSKeyword("RESPONSE")>
    [RESPONSE]

    <FITSKeyword("NOTE")>
    [NOTE]

    <FITSKeyword("TRAKTIME")>
    [TRAKTIME]

    <FITSKeyword("SWMODIFY")>
    [SWMODIFY]

    <FITSKeyword("HISTORY")>
    [HISTORY]

    <FITSKeyword("EXPOSURE")>
    [EXPOSURE]

    <FITSKeyword("CREATOR")>
    [CREATOR]

    <FITSKeyword("CCDXBIN")>
    [CCDXBIN]

    <FITSKeyword("CCDYBIN")>
    [CCDYBIN]



    <FITSKeyword("SCALE")>
    [SCALE]

    <FITSKeyword("PIXSCALE")>
    [PIXSCALE]

    <FITSKeyword("TIME")>
    [TIME]

    <FITSKeyword("BLANK")>
    [BLANK]

    <FITSKeyword("SOFTWARE")>
    [SOFTWARE]


End Enum

'''<summary>Class that provides access to the keyword string and the description of the given keyword.</summary>
Public Class cFITSKey
    Default Public ReadOnly Property Key(ByVal Element As eFITSKeywords) As String()
        Get
            Return FITSKeyword.GetKeywords(Element)
        End Get
    End Property
    Public ReadOnly Property Comment(ByVal Element As eFITSKeywords) As String
        Get
            Return FITSComment.GetComment(Element)
        End Get
    End Property
End Class