Option Explicit On
Option Strict On

'================================================================================
'''<summary>Attribute used for hold the FITS keyword.</summary>
<AttributeUsage(AttributeTargets.All, Inherited:=True, AllowMultiple:=False)>
Public Class FITSKeyword
    Inherits System.Attribute
    Public Sub New()
        Me.New(String.Empty)
    End Sub
    Public Sub New(ByVal Keyword As String)
        MyKeyword = Keyword
    End Sub
    Public ReadOnly Property Keyword() As String
        Get
            Return MyKeyword
        End Get
    End Property
    Private MyKeyword As String = String.Empty
    '''<summary>The keyword string associated with the given keyword.</summary>
    Public Shared Function GetKeyword(ByRef Element As eFITSKeywords) As String
        Dim attributes As FITSKeyword() = CType(Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(FITSKeyword), False), FITSKeyword())
        If attributes.Length > 0 Then Return attributes(0).Keyword Else Return String.Empty
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

'''<summary>Class to provide FITS header keywords (elements and service functions).</summary>
'''<see cref="http://wise2.ipac.caltech.edu/docs/release/prelim/expsup/sec2_3b.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/standard_dict.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
'''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
'''<see cref="http://eso-python.github.io/ESOPythonTutorials/FITS-images.html"/>
Public Enum eFITSKeywords

    '''<summary>The value field shall contain a character string identifying who compiled the information In the data associated With the header. This keyword Is appropriate When the data originate In a published paper Or are compiled from many sources.</summary>
    <FITSKeyword("AUTHOR")>
    <FITSComment("")>
    [AUTHOR]

    '''<summary>If present the image has a valid Bayer color pattern.</summary>
    <FITSKeyword("BAYERPAT")>
    <FITSComment("")>
    [BAYERPAT]

    '''<summary>8 unsigned int, 16 & 32 int, -32 & -64 real.</summary>
    <FITSKeyword("BITPIX")>
    <FITSComment("8 unsigned int, 16 & 32 int, -32 & -64 real")>
    [BITPIX]

    '''<summary>Configured BRIGHTNESS value of the camera.</summary>
    <FITSKeyword("BRIGHTN")>
    <FITSComment("")>
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

    '''<summary>Type of color sensor Bayer array or zero for monochrome.</summary>
    <FITSKeyword("COLORTYP")>
    <FITSComment("")>
    [COLORTYP]

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
    <FITSComment("")>
    [CRVAL1]

    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    '''<remarks> Declination at CRPIX1,CRPIX2 for EQUINOX.</remarks>
    <FITSKeyword("CRVAL2")>
    <FITSComment("")>
    [CRVAL2]

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002</summary>
    '''<example>'RA---SIN'</example>
    <FITSKeyword("CTYPE1")>
    <FITSComment("Projection type for axis 1.")>
    [CTYPE1]

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002</summary>
    '''<example>'DEC--SIN'</example>
    <FITSKeyword("CTYPE2")>
    <FITSComment("Projection type for axis 2.")>
    [CTYPE2]

    '''<summary>Used to color encoding.</summary>
    <FITSKeyword("[CTYPE3]")>
    <FITSComment("Used to color encoding.")>
    [CTYPE3]

    '''<summary>The value field shall contain a character string that gives the date on which the observation ended, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword("DATE_END")>
    <FITSComment("")>
    [DATE_END]

    '''<summary>The value field shall contain a character string that gives the date on which the observation started, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword("DATE_OBS")>
    <FITSComment("observation start, UT")>
    [DATE_OBS]

    '''<summary>The value field gives the declination of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'dd:mm:ss.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <FITSKeyword("DEC")>
    <FITSComment("")>
    [DEC]

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

    '''<summary>Equinox of the World Coordinate System (WCS).</summary>
    '''<example>2000.0</example>
    <FITSKeyword("EQUINOX")>
    <FITSComment("Equinox of the World Coordinate System (WCS)")>
    [EQUINOX]

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

    '''<summary>The value field shall contain a floating point number giving the geographic latitude from which the observation was made in units of degrees.</summary>
    <FITSKeyword("LATITUDE")>
    <FITSComment("Geographic latitude from which the observation was made - degree")>
    [LATITUDE]

    '''<summary>Primary HDU - Number of data axes. Always = 2 for two-dimensional images.</summary>
    <FITSKeyword("NAXIS")>
    <FITSComment("number of axes")>
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
    <FITSComment("Length of data axis 3 or number of color channels")>
    [NAXIS3]

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
    <FITSKeyword("OBS_ID")>
    <FITSComment("")>
    [OBS_ID]

    '''<summary>Configured OFFSET value of the camera.</summary>
    <FITSKeyword("OFFSET")>
    <FITSComment("Configured OFFSET value of the camera")>
    [OFFSET]

    '''<summary>The value field shall contain a character string identifying the organization or institution responsible for creating the FITS file.</summary>
    <FITSKeyword("ORIGIN")>
    <FITSComment("")>
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
    <FITSComment("")>
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

    '''<summary>The value field gives the Right Ascension of the observation.  It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'HH:MM:SS.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <FITSKeyword("RA")>
    <FITSComment("")>
    [RA]

    '''<summary>CCD temperature setpoint in degrees C. Absent if setpoint was not entered.</summary>
    <FITSKeyword("SET-TEMP")>
    <FITSComment("CCD temperature setpoint in C")>
    [SETTEMP]

    '''<summary>Primary HDU.</summary>
    <FITSKeyword("SIMPLE")>
    <FITSComment("")>
    [SIMPLE]

    '''<summary>Clear aperture of the telescope [m].</summary>
    <FITSKeyword("TELAPER")>
    <FITSComment("")>
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
    <FITSKeyword("TIME_END")>
    <FITSComment("")>
    [TIME_END]

    ''''<summary>The value field shall contain a character string that gives the time at which the observation started, format 'hh:mm:ss.sss'.</summary>
    <FITSKeyword("TIME_OBS")>
    <FITSComment("")>
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

End Enum

'''<summary>Class that provides access to the keyword string and the description of the given keyword.</summary>
Public Class cFITSKey
    Default Public ReadOnly Property Key(ByVal Element As eFITSKeywords) As String
        Get
            Return FITSKeyword.GetKeyword(Element)
        End Get
    End Property
    Public ReadOnly Property Comment(ByVal Element As eFITSKeywords) As String
        Get
            Return FITSComment.GetComment(Element)
        End Get
    End Property
End Class

Public Class cFITSKeywords

    '''<summary>Describes which data type a certain keyword has.</summary>
    Public Shared Function GetDataType(ByVal Keyword As eFITSKeywords) As String
        'Set the correct data type
        Select Case Keyword
            Case eFITSKeywords.BITPIX, eFITSKeywords.NAXIS, eFITSKeywords.NAXIS1, eFITSKeywords.NAXIS2
                Return "INTEGER"
            Case eFITSKeywords.BZERO, eFITSKeywords.BSCALE
                Return "DOUBLE"
            Case Else
                Return String.Empty
        End Select
    End Function

    Public Shared Function AsString(ByVal Value As Object) As String
        Dim TypeName As String = Value.GetType.Name
        Select Case TypeName
            Case "String"
                Return cFITSKeywords.GetString(CStr(Value))
            Case "Double"
                Return cFITSKeywords.GetDouble(CDbl(Value))
            Case "Int32"
                Return CStr(Value).Trim
            Case Else
                Return CStr(Value)
        End Select
    End Function

    '''<summary>Formated content as a string.</summary>
    Public Shared Function GetString(ByVal Value As String) As String
        Return "'" & Value & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, without time.</summary>
    Public Shared Function GetDouble(ByVal Value As Double) As String
        Return Value.ToString.Trim.Replace(",", ".")
    End Function

    '''<summary>Formated content for all "DATE..." fields, without time.</summary>
    Public Shared Function GetDate() As String
        Return "'" & GetDate(Now) & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, without time.</summary>
    Public Shared Function GetDate(ByVal Moment As DateTime) As String
        Return "'" & Format(Moment, "yyyy-dd-MM") & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, time.</summary>
    Public Shared Function GetDateWithTime() As String
        Return "'" & GetDateWithTime(Now) & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, time.</summary>
    Public Shared Function GetDateWithTime(ByVal Moment As DateTime) As String
        Return "'" & Format(Moment, "yyyy-dd-MMTHH:mm:ss.fff") & "'"
    End Function

    '''<summary>Formated content for all "TIME..." fields, time.</summary>
    Public Shared Function GetTime(ByVal Moment As DateTime) As String
        Return "'" & Format(Moment, "HH:mm:ss.fff") & "'"
    End Function

End Class