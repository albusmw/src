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
    <ComponentModel.Description("")>
    [AUTHOR]

    '''<summary>If present the image has a valid Bayer color pattern.</summary>
    <FITSKeyword("BAYERPAT")>
    <ComponentModel.Description("")>
    [BAYERPAT]

    '''<summary>8 unsigned int, 16 & 32 int, -32 & -64 real.</summary>
    <FITSKeyword("BITPIX")>
    <FITSComment("8 unsigned int, 16 & 32 int, -32 & -64 real")>
    <ComponentModel.Description("8 unsigned int, 16 & 32 int, -32 & -64 real")>
    [BITPIX]

    '''<summary>Configured BRIGHTNESS value of the camera.</summary>
    <FITSKeyword("BRIGHTN")>
    <ComponentModel.Description("")>
    [BRIGHTNESS]

    '''<summary>Zero point in scaling equation.</summary>
    <FITSKeyword("BZERO")>
    <ComponentModel.Description("")>
    [BZERO]

    '''<summary>Actual measured sensor temperature at the start of exposure in degrees C. Absent if temperature is not available.</summary>
    <FITSKeyword("CCD-TEMP")>
    <ComponentModel.Description("")>
    [CCDTEMP]

    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 1 pixel scale at CRPIX1,CRPIX2. See PXSCAL1 For arcsec equivalent.</remarks>
    '''<example>-0.0003819444391411 degrees/pix</example>
    <FITSKeyword("CDELT1")>
    <ComponentModel.Description("Axis 1 pixel scale at CRPIX1,CRPIX2")>
    [CDELT1]

    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 2 pixel scale at CRPIX1,CRPIX2. See PXSCAL2 For arcsec equivalent.</remarks>
    '''<example>0.0003819444391411 degrees/pix</example>
    <FITSKeyword("CDELT2")>
    <ComponentModel.Description("Axis 2 pixel scale at CRPIX1,CRPIX2")>
    [CDELT2]

    '''<summary>Type of color sensor Bayer array or zero for monochrome.</summary>
    <FITSKeyword("COLORTYP")>
    <ComponentModel.Description("")>
    [COLORTYP]

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
    <ComponentModel.Description("")>
    [CRVAL1]

    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    '''<remarks> Declination at CRPIX1,CRPIX2 for EQUINOX.</remarks>
    <FITSKeyword("CRVAL2")>
    <ComponentModel.Description("")>
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

    '''<summary>The value field shall contain a character string that gives the date on which the observation ended, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword("DATE_END")>
    <ComponentModel.Description("")>
    [DATE_END]

    '''<summary>The value field shall contain a character string that gives the date on which the observation started, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <FITSKeyword("DATE_OBS")>
    <ComponentModel.Description("YYYY-MM-DDThh:mm:ss observation start, UT")>
    [DATE_OBS]

    '''<summary>The value field gives the declination of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'dd:mm:ss.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <FITSKeyword("DEC")>
    <ComponentModel.Description("")>
    [DEC]

    '''<summary>Primary HDU.</summary>
    <FITSKeyword("END")>
    <ComponentModel.Description("")>
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
    <FITSComment("Exposure time in seconds")>
    [FOCUS]

    '''<summary>Field of view [°] along axis 1.</summary>
    <FITSKeyword("FOV1")>
    <ComponentModel.Description("")>
    [FOV1]

    '''<summary>Field of view [°] along axis 2.</summary>
    <FITSKeyword("FOV2")>
    <ComponentModel.Description("")>
    [FOV2]

    '''<summary>Configured GAIN value of the camera.</summary>
    <FITSKeyword("GAIN")>
    <ComponentModel.Description("")>
    [GAIN]

    '''<summary>Type of image: Light Frame, Bias Frame, Dark Frame, Flat Frame, or Tricolor Image.</summary>
    <FITSKeyword("IMAGETYP")>
    <ComponentModel.Description("")>
    [IMAGETYP]

    '''<summary>The value field shall contain a character string identifying the instrument used to acquire the data associated with the header.</summary>
    <FITSKeyword("INSTRUME")>
    <ComponentModel.Description("")>
    [INSTRUME]

    '''<summary>Primary HDU - Number of data axes. Always = 2 for two-dimensional images.</summary>
    <FITSKeyword("NAXIS")>
    <ComponentModel.Description("")>
    [NAXIS]

    '''<summary>Primary HDU - Length of data axis 1 or number of columns in image.</summary>
    <FITSKeyword("NAXIS1")>
    <ComponentModel.Description("")>
    [NAXIS1]

    '''<summary>Primary HDU - Length of data axis 2 or number of rows in image.</summary>
    <FITSKeyword("NAXIS2")>
    <ComponentModel.Description("")>
    [NAXIS2]

    '''<summary>Primary HDU - Length of data axis 3 or number of color channels in image.</summary>
    <FITSKeyword("NAXIS3")>
    <ComponentModel.Description("")>
    [NAXIS3]

    '''<summary>The value field shall contain a character string giving a name for the object observed.</summary>
    <FITSKeyword("OBJECT")>
    <ComponentModel.Description("")>
    [OBJECT]

    '''<summary>The value field shall contain a character string giving a name for the observed object that conforms to the IAU astronomical Object naming conventions. The value of this keyword Is more strictly constrained than For the standard Object keyword which In practice has often been used To record other ancillary information about the observation (e.g. filter, exposure time, weather conditions, etc.).</summary>
    <FITSKeyword("OBJNAME")>
    <ComponentModel.Description("")>
    [OBJNAME]

    '''<summary>The value field shall contain a character string identifying who acquired the data associated with the header.</summary>
    <FITSKeyword("OBSERVER")>
    <ComponentModel.Description("")>
    [OBSERVER]

    ''''<summary>The value field shall contain a character string which uniquely identifies the dataset contained In the FITS file. This Is typically a sequence number that can contain a mixture Of numerical And character values. Example '10315-01-01-30A'.</summary>
    <FITSKeyword("OBS_ID")>
    <ComponentModel.Description("")>
    [OBS_ID]

    '''<summary>Configured OFFSET value of the camera.</summary>
    <FITSKeyword("OFFSET")>
    <ComponentModel.Description("Configured OFFSET value of the camera")>
    [OFFSET]

    '''<summary>The value field shall contain a character string identifying the organization or institution responsible for creating the FITS file.</summary>
    <FITSKeyword("ORIGIN")>
    <ComponentModel.Description("")>
    [ORIGIN]

    '''<summary>Pixel size [um] along axis 1.</summary>
    <FITSKeyword("PIXSIZE1")>
    <ComponentModel.Description("Pixel size [um] along axis 1")>
    [PIXSIZE1]

    '''<summary>Pixel size [um] along axis 2.</summary>
    <FITSKeyword("PIXSIZE2")>
    <ComponentModel.Description("Pixel size [um] along axis 2")>
    [PIXSIZE2]

    ''<summary>Plate size [cm] along axis 1.</summary>
    <FITSKeyword("PLATESZ1]")>
    <ComponentModel.Description("Plate size [cm] along axis 1")>
    [PLATESZ1]

    ''<summary>Plate size [cm] along axis 2.</summary>
    <FITSKeyword("PLATESZ2]")>
    <ComponentModel.Description("Plate size [cm] along axis 2")>
    [PLATESZ2]

    '''<summary>The value field shall contain a character string giving the name, And optionally, the version of the program that originally created the current FITS HDU. This keyword Is synonymous With the CREATOR keyword.  Example 'TASKNAME V1.2.3'.</summary>
    <FITSKeyword("PROGRAM")>
    <ComponentModel.Description("")>
    [PROGRAM]

    '''<summary>Pixel scale at CRPIX1,CRPIX2 for axis1.</summary>
    '''<example>1.375 arcsec/pix</example>
    <FITSKeyword("PXSCAL1")>
    <ComponentModel.Description("Pixel scale at CRPIX1,CRPIX2 for axis1.")>
    [PXSCAL1]

    '''<summary>Pixel scale at CRPIX1,CRPIX2 for axis2.</summary>
    '''<example>1.375 arcsec/pix</example>
    <FITSKeyword("PXSCAL21")>
    <ComponentModel.Description("Pixel scale at CRPIX1,CRPIX2 for axis1.")>
    [PXSCAL2]

    '''<summary>QHY read-out mode.</summary>
    <FITSKeyword("QHY_MODE")>
    <ComponentModel.Description("QHY read-out mode")>
    [QHY_MODE]

    '''<summary>The value field gives the Right Ascension of the observation.  It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'HH:MM:SS.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <FITSKeyword("RA")>
    <ComponentModel.Description("")>
    [RA]

    '''<summary>CCD temperature setpoint in degrees C. Absent if setpoint was not entered.</summary>
    <FITSKeyword("SET-TEMP")>
    <ComponentModel.Description("")>
    [SETTEMP]

    '''<summary>Primary HDU.</summary>
    <FITSKeyword("SIMPLE")>
    <ComponentModel.Description("")>
    [SIMPLE]

    '''<summary>Clear aperture of the telescope [m].</summary>
    <FITSKeyword("TELAPER")>
    <ComponentModel.Description("")>
    [TELAPER]

    '''<summary>The value field shall contain a character string identifying the telescope used to acquire the data associated with the header.</summary>
    <FITSKeyword("TELESCOP")>
    <ComponentModel.Description("Telescope used to acquire data.")>
    [TELESCOP]

    ''<summary>Focal length of the telescope [m].</summary>
    <FITSKeyword("TELFOC")>
    <ComponentModel.Description("Focal length of the telescope [m]")>
    [TELFOC]

    '''<summary>Plate scale of the telescope [arcsec/mm].</summary>
    <FITSKeyword("TELSCALE")>
    <ComponentModel.Description("Plate scale of the telescope [arcsec/mm]")>
    [TELSCALE]

    '''<summary>The value field shall contain a character string that gives the time at which the observation ended, format 'hh:mm:ss.sss'.</summary>
    <FITSKeyword("TIME_END")>
    <ComponentModel.Description("")>
    [TIME_END]

    ''''<summary>The value field shall contain a character string that gives the time at which the observation started, format 'hh:mm:ss.sss'.</summary>
    <FITSKeyword("TIME_OBS")>
    <ComponentModel.Description("")>
    [TIME_OBS]

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


Public Structure sFITSKeywords



    '''<summary>Focuser temperature readout in degrees C, if available.</summary>
    Public Const [FOCUSTEM] As String = "FOCUSTEM"

    '''<summary>Electronic gain in photoelectrons per ADU.</summary>
    Public Const [EGAIN] As String = "EGAIN"

    '''<summary>The value field shall contain a floating point number giving the geographic latitude from which the observation was made in units of degrees.</summary>
    Public Const [LATITUDE] As String = "LATITUDE"

    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    Public Const [CROTA1] As String = "CROTA1"
    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    Public Const [CROTA2] As String = "CROTA2"


    '''<summary>Used to color encoding.</summary>
    Public Const [CTYPE3] As String = "CTYPE3"


    '''<summary>Configured OFFSET value of the camera.</summary>
    Public Shared Function GetComment(ByRef Element As Object) As String
        Dim attributes As Object() = Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(ComponentModel.DescriptionAttribute), False)
        If attributes.Length > 0 Then
            Return "Test"
        Else
            Return String.Empty
        End If
    End Function

End Structure

Public Class cFITSKeywords

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