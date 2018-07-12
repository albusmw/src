Option Explicit On
Option Strict On

'''<summary>Class to provide FITS header keywords (elements and service functions).</summary>
'''<remarks>See e.g. https://heasarc.gsfc.nasa.gov/docs/fcg/common_dict.html. </remarks>
Public Structure eFITSKeywords

    Public Const [SIMPLE] As String = "SIMPLE"          'Primary HDU
    Public Const [BITPIX] As String = "BITPIX"          'Primary HDU
    Public Const [NAXIS] As String = "NAXIS"            'Primary HDU
    Public Const [NAXIS1] As String = "NAXIS1"          'Primary HDU
    Public Const [END] As String = "END"                'Primary HDU

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

    '=============================================================================
    'Found in FITS and need additional comments

    '''<summary>Used to color encoding.</summary>
    Public Const [CTYPE3] As String = "CTYPE3"

End Structure

Public Class cFITSKeywords

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