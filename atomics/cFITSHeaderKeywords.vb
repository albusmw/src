Option Explicit On
Option Strict On

'================================================================================
'''<summary>Attribute used for hold the FITS keyword.</summary>
<AttributeUsage(AttributeTargets.All, Inherited:=True, AllowMultiple:=False)>
Public Class Keyword
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
    Public Shared Function GetKeyword(ByRef Element As eFITSKeywords) As String
        Dim attributes As Keyword() = CType(Element.GetType.GetField(Element.ToString).GetCustomAttributes(GetType(Keyword), False), Keyword())
        If attributes.Length > 0 Then Return attributes(0).Keyword Else Return String.Empty
    End Function
End Class
'================================================================================

'''<summary>Class to provide FITS header keywords (elements and service functions).</summary>
'''<see cref="http://wise2.ipac.caltech.edu/docs/release/prelim/expsup/sec2_3b.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/standard_dict.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
'''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
Public Enum eFITSKeywords


    '''<summary>The value field shall contain a character string identifying who compiled the information In the data associated With the header. This keyword Is appropriate When the data originate In a published paper Or are compiled from many sources.</summary>
    <Keyword("AUTHOR")>
    <ComponentModel.Description("")>
    [AUTHOR]

    '''<summary>If present the image has a valid Bayer color pattern.</summary>
    <Keyword("BAYERPAT")>
    <ComponentModel.Description("")>
    [BAYERPAT]

    '''<summary>8 unsigned int, 16 & 32 int, -32 & -64 real.</summary>
    <Keyword("BITPIX")>
    <ComponentModel.Description("8 unsigned int, 16 & 32 int, -32 & -64 real")>
    [BITPIX]

    '''<summary>Configured BRIGHTNESS value of the camera.</summary>
    <Keyword("BRIGHTN")>
    <ComponentModel.Description("")>
    [BRIGHTNESS]

    '''<summary>Zero point in scaling equation.</summary>
    <Keyword("BZERO")>
    <ComponentModel.Description("")>
    [BZERO]

    '''<summary>Actual measured sensor temperature at the start of exposure in degrees C. Absent if temperature is not available.</summary>
    <Keyword("CCD-TEMP")>
    <ComponentModel.Description("")>
    [CCDTEMP]

    '''<summary>Type of color sensor Bayer array or zero for monochrome.</summary>
    <Keyword("COLORTYP")>
    <ComponentModel.Description("")>
    [COLORTYP]

    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS1+1)</remarks>
    <Keyword("CRPIX1")>
    <ComponentModel.Description("")>
    [CRPIX1]

    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS2+1)</remarks>
    <Keyword("CRPIX2")>
    <ComponentModel.Description("")>
    [CRPIX2]

    '''<summary>The value field shall contain a character string that gives the date on which the observation ended, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <Keyword("DATE_END")>
    <ComponentModel.Description("")>
    [DATE_END]

    '''<summary>The value field shall contain a character string that gives the date on which the observation started, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    <Keyword("DATE_OBS")>
    <ComponentModel.Description("YYYY-MM-DDThh:mm:ss observation start, UT")>
    [DATE_OBS]

    '''<summary>The value field gives the declination of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'dd:mm:ss.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <Keyword("DEC")>
    <ComponentModel.Description("")>
    [DEC]

    '''<summary>Primary HDU.</summary>
    <Keyword("END")>
    <ComponentModel.Description("")>
    [END]

    '''<summary>The value field shall contain a floating point number giving the exposure time of the observation in units of seconds.</summary>
    <Keyword("EXPTIME")>
    <ComponentModel.Description("Exposure time in seconds")>
    [EXPTIME]

    '''<summary>Field of view [°] along axis 1.</summary>
    <Keyword("FOV1")>
    <ComponentModel.Description("")>
    [FOV1]

    '''<summary>Field of view [°] along axis 2.</summary>
    <Keyword("FOV2")>
    <ComponentModel.Description("")>
    [FOV2]

    '''<summary>Configured GAIN value of the camera.</summary>
    <Keyword("GAIN")>
    <ComponentModel.Description("")>
    [GAIN]

    '''<summary>Type of image: Light Frame, Bias Frame, Dark Frame, Flat Frame, or Tricolor Image.</summary>
    <Keyword("IMAGETYP")>
    <ComponentModel.Description("")>
    [IMAGETYP]

    '''<summary>The value field shall contain a character string identifying the instrument used to acquire the data associated with the header.</summary>
    <Keyword("INSTRUME")>
    <ComponentModel.Description("")>
    [INSTRUME]

    '''<summary>Primary HDU.</summary>
    <Keyword("NAXIS")>
    <ComponentModel.Description("")>
    [NAXIS]

    '''<summary>Primary HDU.</summary>
    <Keyword("NAXIS1")>
    <ComponentModel.Description("")>
    [NAXIS1]

    '''<summary>The value field shall contain a character string giving a name for the object observed.</summary>
    <Keyword("OBJECT")>
    <ComponentModel.Description("")>
    [OBJECT]

    '''<summary>The value field shall contain a character string giving a name for the observed object that conforms to the IAU astronomical Object naming conventions. The value of this keyword Is more strictly constrained than For the standard Object keyword which In practice has often been used To record other ancillary information about the observation (e.g. filter, exposure time, weather conditions, etc.).</summary>
    <Keyword("OBJNAME")>
    <ComponentModel.Description("")>
    [OBJNAME]

    ''''<summary>The value field shall contain a character string which uniquely identifies the dataset contained In the FITS file. This Is typically a sequence number that can contain a mixture Of numerical And character values. Example '10315-01-01-30A'.</summary>
    <Keyword("OBS_ID")>
    <ComponentModel.Description("")>
    [OBS_ID]

    '''<summary>Configured OFFSET value of the camera.</summary>
    <Keyword("OFFSET")>
    <ComponentModel.Description("")>
    [OFFSET]

    '''<summary>The value field shall contain a character string identifying the organization or institution responsible for creating the FITS file.</summary>
    <Keyword("ORIGIN")>
    <ComponentModel.Description("")>
    [ORIGIN]

    '''<summary>Pixel size [um] along axis 1.</summary>
    <Keyword("PIXSIZE1")>
    <ComponentModel.Description("")>
    [PIXSIZE1]

    '''<summary>Pixel size [um] along axis 2.</summary>
    <Keyword("PIXSIZE2")>
    <ComponentModel.Description("")>
    [PIXSIZE2]

    ''<summary>Plate size [cm] along axis 1.</summary>
    <Keyword("[PLATESZ1]")>
    <ComponentModel.Description("")>
    [PLATESZ1]

    ''<summary>Plate size [cm] along axis 2.</summary>
    <Keyword("[PLATESZ2]")>
    <ComponentModel.Description("")>
    [PLATESZ2]

    '''<summary>The value field shall contain a character string giving the name, And optionally, the version of the program that originally created the current FITS HDU. This keyword Is synonymous With the CREATOR keyword.  Example 'TASKNAME V1.2.3'.</summary>
    <Keyword("PROGRAM")>
    <ComponentModel.Description("")>
    [PROGRAM]

    '''<summary>QHY read-out mode.</summary>
    <Keyword("QHY_MODE")>
    <ComponentModel.Description("")>
    [QHY_MODE]

    '''<summary>The value field gives the Right Ascension of the observation.  It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'HH:MM:SS.sss' format where the decimal point and number of fractional digits are optional.</summary>
    <Keyword("RA")>
    <ComponentModel.Description("")>
    [RA]

    '''<summary>CCD temperature setpoint in degrees C. Absent if setpoint was not entered.</summary>
    <Keyword("SET-TEMP")>
    <ComponentModel.Description("")>
    [SETTEMP]

    '''<summary>Primary HDU.</summary>
    <Keyword("SIMPLE")>
    <ComponentModel.Description("")>
    [SIMPLE]

    '''<summary>Clear aperture of the telescope [m].</summary>
    <Keyword("TELAPER")>
    <ComponentModel.Description("")>
    [TELAPER]

    '''<summary>The value field shall contain a character string identifying the telescope used to acquire the data associated with the header.</summary>
    <Keyword("TELESCOP")>
    <ComponentModel.Description("")>
    [TELESCOP]

    ''<summary>Focal length of the telescope [m].</summary>
    <Keyword("TELFOC")>
    <ComponentModel.Description("")>
    [TELFOC]

    '''<summary>Plate scale of the telescope [arcsec/mm].</summary>
    <Keyword("TELSCALE")>
    <ComponentModel.Description("")>
    [TELSCALE]

    '''<summary>The value field shall contain a character string that gives the time at which the observation ended, format 'hh:mm:ss.sss'.</summary>
    <Keyword("TIME_END")>
    <ComponentModel.Description("")>
    [TIME_END]

    ''''<summary>The value field shall contain a character string that gives the time at which the observation started, format 'hh:mm:ss.sss'.</summary>
    <Keyword("TIME_OBS")>
    <ComponentModel.Description("")>
    [TIME_OBS]

End Enum

'''<summary>Class that provides access to the keyword string and the description of the given keyword.</summary>
Public Class cFITSKey
    Default Public ReadOnly Property Key(ByVal Element As eFITSKeywords) As String
        Get
            Return Keyword.GetKeyword(Element)
        End Get
    End Property
End Class


Public Structure sFITSKeywords


    '''<summary>The value field shall contain a character string identifying who acquired the data associated with the header.</summary>
    Public Const [OBSERVER] As String = "OBSERVER"

    '''<summary>Focus value (from logbook). Used when a single value is given in the logs.</summary>
    Public Const [FOCUS] As String = "FOCUS"



    '''<summary>Focuser temperature readout in degrees C, if available.</summary>
    Public Const [FOCUSTEM] As String = "FOCUSTEM"

    '''<summary>Electronic gain in photoelectrons per ADU.</summary>
    Public Const [EGAIN] As String = "EGAIN"

    '''<summary>The value field shall contain a floating point number giving the geographic latitude from which the observation was made in units of degrees.</summary>
    Public Const [LATITUDE] As String = "LATITUDE"


    '=============================================================================
    'Found in FITS and need additional comments

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002.</summary>
    Public Const [CTYPE1] As String = "CTYPE1"
    '''<summary>Projection type for axis 2. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002.</summary>
    Public Const [CTYPE2] As String = "CTYPE2"

    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 1 pixel scale at CRPIX1,CRPIX2. See PXSCAL1 For arcsec equivalent.</remarks>
    Public Const [CDELT1] As String = "CDELT1"
    '''<summary>The value field shall contain a floating point number giving the Partial derivative Of the coordinate specified by the CTYPEn keywords with respect to the pixel index, evaluated at the reference point CRPIXn, in units Of the coordinate specified by  the CTYPEn keyword.  These units must follow the prescriptions of section 5.3 of the FITS Standard.</summary>
    '''<remarks>Axis 1 pixel scale at CRPIX1,CRPIX2. See PXSCAL1 For arcsec equivalent.</remarks>
    Public Const [CDELT2] As String = "CDELT2"

    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    Public Const [CROTA1] As String = "CROTA1"
    '''<summary>This keyword is used to indicate a rotation from a standard coordinate system described by the CTYPEn To a different coordinate system in which the values in the array are actually expressed. Rules For such rotations are Not further specified in the Standard; the rotation should be explained In comments. The value field shall contain a floating point number giving the rotation angle In degrees between axis n And the direction implied by the coordinate system defined by CTYPEn.</summary>
    '''<remarks>UNITS: degrees</remarks>
    Public Const [CROTA2] As String = "CROTA2"

    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    Public Const [CRVAL1] As String = "CRVAL1"
    '''<summary>The value field shall contain a floating point number, giving the value Of the coordinate specified by the CTYPEn keyword at the reference point CRPIXn. Units must follow the prescriptions Of section 5.3 of the FITS Standard.</summary>
    Public Const [CRVAL2] As String = "CRVAL2"


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