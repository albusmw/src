Option Explicit On
Option Strict On

'''<summary>Class to provide FITS header keywords (elements and service functions).</summary>
'''<see cref="http://wise2.ipac.caltech.edu/docs/release/prelim/expsup/sec2_3b.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/standard_dict.html"/>
'''<see cref="https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html"/>
'''<see cref="https://diffractionlimited.com/help/maximdl/FITS_File_Header_Definitions.htm"/>
Public Structure eFITSKeywords

    Public Const [SIMPLE] As String = "SIMPLE"          'Primary HDU
    Public Const [BITPIX] As String = "BITPIX"          'Primary HDU
    Public Const [NAXIS] As String = "NAXIS"            'Primary HDU
    Public Const [NAXIS1] As String = "NAXIS1"          'Primary HDU
    Public Const [END] As String = "END"                'Primary HDU

    '''<summary>Zero point in scaling equation.</summary>
    Public Const [BZERO] As String = "BZERO"

    'Alphabetic order ...

    '''<summary>If present the image has a valid Bayer color pattern.</summary>
    Public Const [BAYERPAT] As String = "BAYERPAT"
    '''<summary>Type of color sensor Bayer array or zero for monochrome.</summary>
    Public Const [COLORTYP] As String = "COLORTYP"

    '''<summary>The value field shall contain a character string identifying the organization or institution responsible for creating the FITS file.</summary>
    Public Const [ORIGIN] As String = "ORIGIN"
    '''<summary>The value field shall contain a character string identifying the telescope used to acquire the data associated with the header.</summary>
    Public Const [TELESCOP] As String = "TELESCOP"
    '''<summary>The value field shall contain a character string identifying the instrument used to acquire the data associated with the header.</summary>
    Public Const [INSTRUME] As String = "INSTRUME"
    '''<summary>The value field shall contain a character string identifying who acquired the data associated with the header.</summary>
    Public Const [OBSERVER] As String = "OBSERVER"
    '''<summary>The value field shall contain a character string giving a name for the object observed.</summary>
    Public Const [OBJECT] As String = "OBJECT"
    '''<summary>Clear aperture of the telescope [m].</summary>
    Public Const [TELAPER] As String = "TELAPER"
    '''<summary>Focal length of the telescope [m].</summary>
    Public Const [TELFOC] As String = "TELFOC"
    '''<summary>Plate scale of the telescope [arcsec/mm].</summary>
    Public Const [TELSCALE] As String = "TELFOC"
    '''<summary>Focus value (from logbook). Used when a single value is given in the logs.</summary>
    Public Const [FOCUS] As String = "FOCUS"

    '''<summary>The value field shall contain a character string that gives the date on which the observation started, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    Public Const [DATE_OBS] As String = "DATE-OBS"
    '''<summary>The value field shall contain a character string that gives the date on which the observation ended, format 'yyyy-mm-dd', or 'yyyy-mm-ddThh:mm:ss.sss'.</summary>
    Public Const [DATE_END] As String = "DATE-END"
    '''<summary>The value field shall contain a character string that gives the time at which the observation started, format 'hh:mm:ss.sss'.</summary>
    Public Const [TIME_OBS] As String = "TIME-OBS"
    '''<summary>The value field shall contain a character string that gives the time at which the observation ended, format 'hh:mm:ss.sss'.</summary>
    Public Const [TIME_END] As String = "TIME-END"
    '''<summary>The value field shall contain a floating point number giving the exposure time of the observation in units of seconds.</summary>
    Public Const [EXPTIME] As String = "EXPTIME"
    '''<summary>The value field gives the declination of the observation. It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'dd:mm:ss.sss' format where the decimal point and number of fractional digits are optional.</summary>
    Public Const [DEC] As String = "DEC"
    '''<summary>The value field gives the Right Ascension of the observation.  It may be expressed either as a floating point number in units of decimal degrees, or as a character string in 'HH:MM:SS.sss' format where the decimal point and number of fractional digits are optional.</summary>
    Public Const [RA] As String = "RA"

    '''<summary>Pixel size [um] along axis 1.</summary>
    Public Const [PIXSIZE1] As String = "PIXSIZE1"
    '''<summary>Pixel size [um] along axis 2.</summary>
    Public Const [PIXSIZE2] As String = "PIXSIZE2"
    '''<summary>Plate size [cm] along axis 1.</summary>
    Public Const [PLATESZ1] As String = "PLATESZ1"
    '''<summary>Plate size [cm] along axis 2.</summary>
    Public Const [PLATESZ2] As String = "PLATESZ2"
    '''<summary>Field of view [°] along axis 1.</summary>
    Public Const [FOV1] As String = "FOV1"
    '''<summary>Field of view [°] along axis 1.</summary>
    Public Const [FOV2] As String = "FOV2"

    '''<summary>Actual measured sensor temperature at the start of exposure in degrees C. Absent if temperature is not available.</summary>
    Public Const [CCDTEMP] As String = "CCD-TEMP"
    '''<summary>CCD temperature setpoint in degrees C. Absent if setpoint was not entered.</summary>
    Public Const [SETTEMP] As String = "SET-TEMP"
    '''<summary>Focuser temperature readout in degrees C, if available.</summary>
    Public Const [FOCUSTEM] As String = "FOCUSTEM"
    '''<summary>Type of image: Light Frame, Bias Frame, Dark Frame, Flat Frame, or Tricolor Image.</summary>
    Public Const [IMAGETYP] As String = "IMAGETYP"
    '''<summary>Electronic gain in photoelectrons per ADU.</summary>
    Public Const [EGAIN] As String = "EGAIN"

    '=============================================================================
    'Found in FITS and need additional comments

    '''<summary>Projection type for axis 1. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002.</summary>
    Public Const [CTYPE1] As String = "CTYPE1"
    '''<summary>Projection type for axis 2. Always set to use the SIN (orthographic) projection; For definition, see Calabretta & Greisen, 2002.</summary>
    Public Const [CTYPE2] As String = "CTYPE2"

    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS1+1)</remarks>
    Public Const [CRPIX1] As String = "CRPIX1"
    '''<summary>The value field shall contain a floating point number, identifying the location Of a reference point along axis n, In units Of the axis index.  This value Is based upon a counter that runs from 1 To NAXISn with an increment of 1 per pixel.  The reference point value need Not be that for the center of a pixel nor lie within the actual data array.  Use comments To indicate the location Of the index point relative to the pixel..</summary>
    '''<remarks>For center, set to 0.5*(NAXIS2+1)</remarks>
    Public Const [CRPIX2] As String = "CRPIX2"

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
    '''<summary>Configured GAIN value of the camera.</summary>
    Public Const [GAIN] As String = "GAIN"



End Structure

Public Class cFITSKeywords

    '''<summary>Formated content for all "DATE..." fields, without time..</summary>
    Public Shared Function GetDouble(ByVal Value As Double) As String
        Return Value.ToString.Trim.Replace(",", ".")
    End Function

    '''<summary>Formated content for all "DATE..." fields, without time..</summary>
    Public Shared Function GetDate() As String
        Return "'" & GetDate(Now) & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, without time..</summary>
    Public Shared Function GetDate(ByVal Moment As DateTime) As String
        Return "'" & Format(Moment, "yyyy-dd-MM") & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, time..</summary>
    Public Shared Function GetDateWithTime() As String
        Return "'" & GetDateWithTime(Now) & "'"
    End Function

    '''<summary>Formated content for all "DATE..." fields, time..</summary>
    Public Shared Function GetDateWithTime(ByVal Moment As DateTime) As String
        Return "'" & Format(Moment, "yyyy-dd-MMTHH:mm:ss.fff") & "'"
    End Function

End Class