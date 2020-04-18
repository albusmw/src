Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

  Partial Public Class AstroCalc

    '########################################################################################################################
    ' Known stars
    '########################################################################################################################

    '''<summary>Some well known stars ...</summary>
    Public Class KnownStars
      Public Shared ReadOnly Beteigeuze As New Ato.AstroCalc.sRADec("05h 55m 10,3s", "+07° 24′ 25,4″")
      Public Shared ReadOnly Polaris As New Ato.AstroCalc.sRADec("02h 31m 49.09s", "+89° 15′ 50.8″")
      Public Shared ReadOnly Altair As New Ato.AstroCalc.sRADec("19h 50m 47s", "+08° 52′ 6″")
    End Class

    '''<summary>Julian data 0 is at 1st of january 2000, 00:00:00 UTC.</summary>
    Private Shared ReadOnly J2000_Zero As Double = JulianDateTime(New DateTime(2000, 1, 1, 12, 0, 0, 0, DateTimeKind.Utc))
    '''<summary>Separator sign.</summary>
    Private Shared ReadOnly Sep As Char = CChar("|")

    '########################################################################################################################
    ' Structures
    '########################################################################################################################

    '''<summary>Right Ascension and Declination.</summary>
    Public Structure sRADec
      '''<summary>Right Ascension.</summary>
      Public RA As Double
      '''<summary>Declination.</summary>
      Public DEC As Double
      Public Sub New(ByVal NewRightAscension As String, ByVal NewDeclination As String)
                Me.RA = AstroParser.ParseRA(NewRightAscension)
                Me.DEC = AstroParser.ParseDeclination(NewDeclination)
            End Sub
        End Structure

        '''<summary>Azimut and Altitude.</summary>
        Public Structure sAzAlt
            '''<summary>Azimut.</summary>
            Public AZ As Double
            '''<summary>Altitude.</summary>
            Public ALT As Double
            Public Sub New(ByVal NewAzimut As Double, ByVal NewAltitude As Double)
                Me.AZ = NewAzimut
                Me.ALT = NewAltitude
            End Sub
        End Structure

        '''<summary>Latitude and Longitude.</summary>
        Public Structure sLatLong
            '''<summary>Latitude (N-S).</summary>
            Public Latitude As Double
            '''<summary>Longitude (W-E).</summary>
            Public Longitude As Double
            Public Sub New(ByVal NewLat As Double, ByVal NewLong As Double)
                Me.Latitude = NewLat
                Me.Longitude = NewLong
            End Sub
            Public Sub New(ByVal NewLat As String, ByVal NewLong As String)
                Me.Latitude = AstroParser.ParsePosition(NewLat)
                Me.Longitude = AstroParser.ParsePosition(NewLong)
            End Sub
        End Structure

        '########################################################################################################################
        ' Time and date calculations
        '########################################################################################################################

        Public Shared Function J2000(ByVal DateToCalc As DateTime) As Double
            Return JulianDateTime(DateToCalc) - J2000_Zero
        End Function

        '''<summary>Calculate the julian date.</summary>
        '''<param name="TC">Date to convert.</param>
        '''<returns>Julian date at .</returns>
        '''<remarks>Formular taken from [Montenbruck], page 42.</remarks>
        Public Shared Function JulianDate(ByVal TC As DateTime) As Double

            Dim y As Integer
            Dim m As Integer
            Dim B As Integer

            If TC.Month <= 2 Then
                y = TC.Year - 1
                m = TC.Month + 12
            Else
                y = TC.Year
                m = TC.Month
            End If

            If TC <= New Date(1582, 10, 4) Then
                B = -2
            Else
                B = FloorToInt(y / 400) - FloorToInt(y / 100)
            End If

            Return FloorToInt(365.25 * y) + FloorToInt(30.6001 * (m + 1)) + B + 1720996.5 + TC.Day

        End Function

        '''<summary>Calculate the julian date and time.</summary>
        '''<param name="TC">Date to convert.</param>
        '''<returns>Julian date and time.</returns>
        '''<remarks>Formular taken from [Montenbruck], page 42.</remarks>
        Public Shared Function JulianDateTime(ByVal TC As DateTime) As Double

            Dim y As Integer
            Dim m As Integer
            Dim B As Integer

            If TC.Month <= 2 Then
                y = TC.Year - 1
                m = TC.Month + 12
            Else
                y = TC.Year
                m = TC.Month
            End If
            If TC <= New Date(1582, 10, 4) Then
                B = -2
            Else
                B = FloorToInt(y / 400) - FloorToInt(y / 100)
            End If

            Dim DateOnly As Double = FloorToInt(365.25 * y) + FloorToInt(30.6001 * (m + 1)) + B + 1720996.5 + TC.Day  'date at midnight
            Return DateOnly + (SecondsSinceMidnight(TC.ToUniversalTime) / 86400)                                      'TODO: UTC is not correct here ...

        End Function

        '''<summary>Calculate the Greenwich Mean Sidereal Time.</summary>
        '''<returns>Greenwich mean sidereal time - date time kind is UTC</returns>
        '''<remarks>Year, day and time are the current year, day and time as UTC value - see also here: "http://www.cv.nrao.edu/~rfisher/Ephemerides/times.html".</remarks>
        Public Shared Function GMST(ByVal LocalDateTime As DateTime) As DateTime

            Dim JD As Double = JulianDate(LocalDateTime)
            Dim T As Double = (JD - 2451545.0) / 36525      '2451545 = 1.1.2000

            Dim Seconds As Double = 24110.54841 + (8640184.812866 * T) + (0.093104 * (T ^ 2)) + (0.0000062 * (T ^ 3))

            Dim RealHours As Integer = FloorToInt(Seconds / 3600) : Seconds -= (RealHours * 3600)
            Dim RealMinutes As Integer = FloorToInt(Seconds / 60) : Seconds -= (RealMinutes * 60)
            Dim RealSeconds As Integer = FloorToInt(Seconds) : Seconds -= RealSeconds

            RealHours = RealHours Mod 24
            Dim ValueUTC As DateTime = LocalDateTime.ToUniversalTime
            Dim RetVal As New DateTime(ValueUTC.Year, ValueUTC.Month, ValueUTC.Day, 0, 0, 0, 0, DateTimeKind.Utc)

            RetVal = RetVal.AddHours(RealHours)
            RetVal = RetVal.AddMinutes(RealMinutes)
            RetVal = RetVal.AddSeconds(RealSeconds)
            RetVal.AddMilliseconds(Seconds / 1000)

            Return RetVal

        End Function

        '########################################################################################################################
        ' Time and date calculations relative to a location
        '########################################################################################################################

        '''<summary>Calculate the Local Mean Sidereal Time.</summary>
        '''<returns>Local mean sidereal time - date time kind is UTC</returns>
        '''<remarks>Year, day and time are the current year, day and time as UTC value.</remarks>
        Public Shared Function LMST(ByVal Value As DateTime, ByVal Longitude As Double) As DateTime
            Return GMST(Value).AddHours(Longitude / 15)
        End Function

        '''<summary>Calculate the Local Sidereal Time.</summary>
        '''<returns>Local sidereal time [°] - date time kind is UTC.</returns>
        '''<remarks>Formular taken from "http://www.stargazing.net/kepler/altaz.html".</remarks>
        Public Shared Function LST(ByVal Value As DateTime, ByVal Longitude As Double) As Double
            Dim UT_Decimal As Double = (Value.Hour + (Value.Minute / 60) + (Value.Second / 3600))
            Dim RetVal As Double = 100.46 + (0.985647 * J2000(Value)) + Longitude + (15 * UT_Decimal)
            For Rotator As Integer = 1 To 10
                If RetVal >= 360 Then
                    RetVal = RetVal - 360 * (Fix(RetVal / 360))
                Else
                    If RetVal < 0 Then
                        RetVal = RetVal - 360 * (Fix(RetVal / 360))
                    Else
                        Return RetVal
                    End If
                End If
            Next Rotator
            Return RetVal
        End Function

        Public Shared Function LSTFormated(ByVal Value As DateTime, ByVal Longitude As Double) As String
            Dim RetVal As Double = LST(Value, Longitude)
            Return FormatHMS(24 * (RetVal / 360))
        End Function

        '''<summary>Return the Hour angle [°].</summary>
        '''<param name="LST">Local Siderial Time [°].</param>
        '''<param name="RA">Right Ascension [°].</param>
        '''<returns>Hour angle [°].</returns>
        Public Shared Function HourAngle(ByVal LST As Double, ByVal RA As Double) As Double
            If LST < 0 Then LST += 360
            Dim RetVal As Double = LST - RA
            If RetVal < 0 Then RetVal += 360
            Return RetVal
        End Function

        '########################################################################################################################
        ' Position calculation
        '########################################################################################################################

        '''<summary>Basic calculation routine to get the coordinates of an astronomical object.</summary>
        '''<param name="UTCDateTime">Date and time, UTC.</param>
        '''<param name="LatLong">Local coordinates.</param>
        '''<param name="RADec">Galactic coordinates.</param>
        Public Shared Function GetHorizontalPosition(ByVal UTCDateTime As DateTime, ByVal LatLong As sLatLong, ByVal RADec As sRADec) As sAzAlt
            Dim RetVal As sAzAlt
            Dim LST As Double = Ato.AstroCalc.LST(UTCDateTime.ToUniversalTime, LatLong.Longitude)
            RetVal.AZ = Double.NaN
            RetVal.ALT = Double.NaN
            RecDec_To_Horizontal(HourAngle(LST, RADec.RA), RADec.DEC, LatLong.Latitude, RetVal.AZ, RetVal.ALT)
            Return RetVal
        End Function

        '########################################################################################################################
        ' Coordinate transformation
        '########################################################################################################################

        '''<summary>Convert equatorial to horizontal coordinates.</summary>
        '''<remarks>Taken from "http://de.wikipedia.org/wiki/Astronomische_Koordinatensysteme"</remarks>
        Public Shared Sub RecDec_To_Horizontal(ByVal HourAngle As Double, ByVal Declination As Double, ByVal Latitude As Double, ByRef AZ As Double, ByRef Alt As Double)

            Dim DegToRad As Double = Math.PI / 180
            Dim SinDec As Double = Math.Sin(Declination * DegToRad)
            Dim SinLat As Double = Math.Sin(Latitude * DegToRad)
            Dim CosDec As Double = Math.Cos(Declination * DegToRad)
            Dim CosLat As Double = Math.Cos(Latitude * DegToRad)
            Dim CosHA As Double = Math.Cos(HourAngle * DegToRad)
            Dim SinHA As Double = Math.Sin(HourAngle * DegToRad)

            Dim SinAlt As Double = SinDec * SinLat + CosDec * CosLat * CosHA
            Dim CosAlt As Double = Math.Cos(Math.Asin(SinAlt))
            Alt = Math.Asin(SinAlt) / DegToRad

            Dim CosA As Double = (SinDec - SinAlt * SinLat) / (CosAlt * CosLat)
            Dim A As Double = Math.Acos(CosA) / DegToRad
            AZ = A : If SinHA > 0 Then AZ = 360 - AZ

        End Sub

        '########################################################################################################################
        ' Helper functions
        '########################################################################################################################

        '''<summary>Calculate the seconds since midnight.</summary>
        Private Shared Function SecondsSinceMidnight(ByVal ToCalculate As DateTime) As Double
            Return (ToCalculate.Hour * 3600) + (ToCalculate.Minute * 60) + ToCalculate.Second + (ToCalculate.Millisecond / 1000)
        End Function

        Private Shared Function FloorToInt(ByVal Value As Double) As Integer
            Return CInt(Decimal.Floor(CDec(Value)))
        End Function

        '########################################################################################################################
        ' Lunar position calculation
        '########################################################################################################################

        Public Shared Function MoonPosition(ByVal TC As DateTime) As sAzAlt

            'EXCEL from http://www.stargazing.net/kepler/moon3.html
            'Other source: http://idlastro.gsfc.nasa.gov/ftp/pro/astro/moonpos.pro


            'Example
            TC = New DateTime(2016, 8, 26, 12, 0, 0)

            TC = TC.ToUniversalTime

            Dim PerTerm_Long_D As Integer() = {0, 2, 2, 0, 0, 0, 2, 2, 2, 2, 0, 1, 0, 2, 0, 0, 4, 0, 4, 2, 2, 1, 1, 2, 2, 4, 2, 0, 2, 2, 1, 2, 0, 0, 2, 2, 2, 4, 0, 3, 2, 4, 0, 2, 2, 2, 4, 0, 4, 1, 2, 0, 1, 3, 4, 2, 0, 1, 2, 2}
            Dim PerTerm_Long_M As Integer() = {0, 0, 0, 0, 1, 0, 0, -1, 0, -1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, -1, 0, 0, 0, 1, 0, -1, 0, -2, 1, 2, -2, 0, 0, -1, 0, 0, 1, -1, 2, 2, 1, -1, 0, 0, -1, 0, 1, 0, 1, 0, 0, -1, 2, 1, 0, 0}
            Dim PerTerm_Long_M_ As Integer() = {1, -1, 0, 2, 0, 0, -2, -1, 1, 0, -1, 0, 1, 0, 1, 1, -1, 3, -2, -1, 0, -1, 0, 1, 2, 0, -3, -2, -1, -2, 1, 0, 2, 0, -1, 1, 0, -1, 2, -1, 1, -2, -1, -1, -2, 0, 1, 4, 0, -2, 0, 2, 1, -2, -3, 2, 1, -1, 3, -1}
            Dim PerTerm_Long_F As Integer() = {0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, -2, 2, -2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, -2, 2, 0, 2, 0, 0, 0, 0, 0, 0, -2, 0, 0, 0, 0, -2, -2, 0, 0, 0, 0, 0, 0, 0, -2}
            Dim PerTerm_Long_L_coeff As Double() = {6288774, 1274027, 658314, 213618, -185116, -114332, 58793, 57066, 53322, 45758, -40923, -34720, -30383, 15327, -12528, 10980, 10675, 10034, 8548, -7888, -6766, -5163, 4987, 4036, 3994, 3861, 3665, -2689, -2602, 2390, -2348, 2236, -2120, -2069, 2048, -1773, -1595, 1215, -1110, -892, -810, 759, -713, -700, 691, 596, 549, 537, 520, -487, -399, -381, 351, -340, 330, 327, -323, 299, 294, 0}
            Dim PerTerm_Long_R_coeff As Double() = {-20905355, -3699111, -2955968, -569925, 48888, -3149, 246158, -152138, -170733, -204586, -129620, 108743, 104755, 10321, 0, 79661, -34782, -23210, -21636, 24208, 30824, -8379, -16675, -12831, -10445, -11650, 14403, -7003, 0, 10056, 6322, -9884, 5751, 0, -4950, 4130, 0, -3958, 0, 3258, 2616, -1897, -2117, 2354, 0, 0, -1423, -1117, -1571, -1739, 0, -4421, 0, 0, 0, 0, 1165, 0, 0, 8752}

            Dim PerTerm_Lat_D As Integer() = {0, 0, 0, 2, 2, 2, 2, 0, 2, 0, 2, 2, 2, 2, 2, 2, 2, 0, 4, 0, 0, 0, 1, 0, 0, 0, 1, 0, 4, 4, 0, 4, 2, 2, 2, 2, 0, 2, 2, 2, 2, 4, 2, 2, 0, 2, 1, 1, 0, 2, 1, 2, 0, 4, 4, 1, 4, 1, 4, 2}
            Dim PerTerm_Lat_M As Integer() = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 1, -1, -1, -1, 1, 0, 1, 0, 1, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 1, 1, 0, -1, -2, 0, 1, 1, 1, 1, 1, 0, -1, 1, 0, -1, 0, 0, 0, -1, -2}
            Dim PerTerm_Lat_M_ As Integer() = {0, 1, 1, 0, -1, -1, 0, 2, 1, 2, 0, -2, 1, 0, -1, 0, -1, -1, -1, 0, 0, -1, 0, 1, 1, 0, 0, 3, 0, -1, 1, -2, 0, 2, 1, -2, 3, 2, -3, -1, 0, 0, 1, 0, 1, 1, 0, 0, -2, -1, 1, -2, 2, -2, -1, 1, 1, -1, 0, 0}
            Dim PerTerm_Lat_F As Integer() = {1, 1, -1, -1, 1, -1, 1, 1, -1, -1, -1, -1, 1, -1, 1, 1, -1, -1, -1, 1, 3, 1, 1, 1, -1, -1, -1, 1, -1, 1, -3, 1, -3, -1, -1, 1, -1, 1, -1, 1, 1, 1, 1, -1, 3, -1, -1, 1, -1, -1, 1, -1, 1, -1, -1, -1, -1, -1, -1, 1}
            Dim PerTerm_Lat_B_coeff As Double() = {5128122, 280602, 277693, 173237, 55413, 46271, 32573, 17198, 9266, 8822, 8216, 4324, 4200, -3359, 2463, 2211, 2065, -1870, 1828, -1794, -1749, -1565, -1491, -1475, -1410, -1344, -1335, 1107, 1021, 833, 777, 671, 607, 596, 491, -451, 439, 422, 421, -366, -351, 331, 315, 302, -283, -229, 223, 223, -220, -220, -185, 181, -177, 176, 166, -164, 132, -119, 115, 107}

            Dim Days_to_J2000_0 As Double = Math.Round(J2000(TC))
            Dim T As Double = Days_to_J2000_0 / 36525                     '0,166516085
            Dim Tpow2 As Double = T * T
            Dim Tpow3 As Double = T * T * T
            Dim Tpow4 As Double = T * T * T * T

            Dim E As Double = 1 - (0.002516 * T) - (0.0000074 * T * T)    '0,99958084
            Dim E2 As Double = E * E                                      '0,999161856

            'Arguments
            Dim L_ As Double = Modulo(218.3164591 + (481267.88134236 * T) - (0.0013268 * Tpow2) + (Tpow3 / 538841) - (Tpow4 / 65194000), 360)   '77,159799
            Dim D As Double = Modulo(297.8502042 + (445267.1115168 * T) - (0.00163 * Tpow2) + (Tpow3 / 545868) - (Tpow4 / 113065000), 360)      '281,986292
            Dim M As Double = Modulo(357.5291092 + (35999.0502909 * T) - (0.0001536 * Tpow2) + (Tpow3 / 24490000), 360)                         '231,950019
            Dim M_ As Double = Modulo(134.9634114 + (477198.8676313 * T) + (0.008997 * Tpow2) + (Tpow3 / 69699) - (Tpow4 / 14712000), 360)      '36,250805
            Dim F As Double = Modulo(93.2720993 + (483202.0175273 * T) - (0.0034029 * Tpow2) + (Tpow3 / 3526000) - (Tpow4 / 863310000), 360)    '274,180167
            Dim A1 As Double = Modulo(119.75 + 131.849 * T, 360)                                                                                '141,704979
            Dim A2 As Double = Modulo(53.09 + 479264.29 * T, 360)                                                                               '298,303190
            Dim A3 As Double = Modulo(313.45 + 481266.484 * T, 360)                                                                             '172,060696

            Dim L_eccen(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double
            Dim R_eccen(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double
            Dim L_term(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double
            Dim R_term(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double
            Dim B_eccen(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double
            Dim B_term(PerTerm_Long_L_coeff.GetUpperBound(0)) As Double

            Dim L_term_Sum As Double = 0
            Dim R_term_Sum As Double = 0
            Dim B_term_Sum As Double = 0

            For Idx As Integer = 0 To L_eccen.GetUpperBound(0)
                Dim PerTerm_Long_M_abs As Integer = Math.Abs(PerTerm_Long_M(Idx))
                Dim PerTerm_Lat_M_abs As Integer = Math.Abs(PerTerm_Lat_M(Idx))
                L_eccen(Idx) = CDbl(IIf(PerTerm_Long_M_abs = 1, PerTerm_Long_L_coeff(Idx) * E, IIf(PerTerm_Long_M_abs = 2, PerTerm_Long_L_coeff(Idx) * E2, PerTerm_Long_L_coeff(Idx))))
                R_eccen(Idx) = CDbl(IIf(PerTerm_Long_M_abs = 1, PerTerm_Long_R_coeff(Idx) * E, IIf(PerTerm_Long_M_abs = 2, PerTerm_Long_R_coeff(Idx) * E2, PerTerm_Long_R_coeff(Idx))))
                B_eccen(Idx) = CDbl(IIf(PerTerm_Lat_M_abs = 1, PerTerm_Lat_B_coeff(Idx) * E, IIf(PerTerm_Lat_M_abs = 2, PerTerm_Lat_B_coeff(Idx) * E2, PerTerm_Lat_B_coeff(Idx))))
                Dim TrigArg_Long As Double = PerTerm_Long_D(Idx) * ToRad(D) + PerTerm_Long_M(Idx) * ToRad(M) + PerTerm_Long_M_(Idx) * ToRad(M_) + PerTerm_Long_F(Idx) * ToRad(F)
                Dim TrigArg_Lat As Double = PerTerm_Lat_D(Idx) * ToRad(D) + PerTerm_Lat_M(Idx) * ToRad(M) + PerTerm_Lat_M_(Idx) * ToRad(M_) + PerTerm_Lat_F(Idx) * ToRad(F)
                L_term(Idx) = L_eccen(Idx) * Math.Sin(TrigArg_Long)
                R_term(Idx) = R_eccen(Idx) * Math.Cos(TrigArg_Long)
                B_term(Idx) = B_eccen(Idx) * Math.Sin(TrigArg_Lat)
                L_term_Sum += L_term(Idx)                                           '4105557,235
                R_term_Sum += R_term(Idx)                                           '-10978212,33
                B_term_Sum += B_term(Idx)                                           '-5219574,067
            Next Idx

            'Totals and additive terms - longitude and radius vector
            Dim Additional_A1 As Double = 3958 * Math.Sin(ToRad(A1))              '2452,81546
            Dim Additional_L_F As Double = 1962 * Math.Sin(ToRad(L_) - ToRad(F))  '574,3002331
            Dim Additional_A2 As Double = 318 * Math.Sin(ToRad(A2))               '-279,9834032
            Dim FinalTotalLat As Double = L_term_Sum + Additional_A1 + Additional_L_F + Additional_A2 '4108304,367

            'Totals and additive terms - latitude
            Dim Additional_L_ As Double = -2235 * Math.Sin(ToRad(L_))               '-2179,110844
            Dim Additional_A3 As Double = 382 * Math.Sin(ToRad(A3))                 '52,76336022
            Dim Additional_A1_F As Double = 175 * Math.Sin(ToRad(A1) - ToRad(F))    '-129,0747223
            Dim Additional_A1F As Double = 175 * Math.Sin(ToRad(A1) + ToRad(F))     '144,8851177
            Dim Additional_L__M_ As Double = 127 * Math.Sin(ToRad(L_) - ToRad(M_))  '83,16715062
            Dim Additional_L_M_ As Double = 115 * Math.Sin(ToRad(L_) + ToRad(M_))   '105,5333277
            Dim FinalTotalLong As Double = B_term_Sum + Additional_L_ + Additional_A3 + Additional_A1_F + Additional_A1F + Additional_L__M_ + Additional_L_M_   '-5221495,903

            'Results
            Dim Lamda As Double = L_ + (FinalTotalLat / 1000000)                    '81,268103
            Dim Beta As Double = FinalTotalLong / 1000000                           '-5,221496
            Dim Distance As Double = 385000.56 + (R_term_Sum / 1000)                '374022,3
            Dim Pi As Double = Math.Asin(6378.14 / Distance) * (180 / Math.PI)      '0,977103

            Return New sAzAlt(Lamda, Beta)

        End Function

        Private Shared Function ToRad(ByVal Val As Double) As Double
            Return Val * (Math.PI / 180)
        End Function

        Private Shared Function Modulo(ByVal A As Double, ByVal B As Integer) As Double
            Dim A_floor As Integer = CInt(Math.Floor(A) / B)
            Dim RetVal As Double = A - (A_floor * B)
            If RetVal < 0 And A > 0 Then RetVal += B
            Return RetVal
        End Function

        '########################################################################################################################
        ' Formater
        '########################################################################################################################

        Public Shared Function Format360Degree(ByVal Degree As Double) As String
            Return Format360Degree(Degree, 2)
        End Function

        '''<summary>Convert a value between 0 and 24 to 00:00:00 and 23:59:59.</summary>
        Public Shared Function FormatHMS(ByVal Value As Double) As String
            Return FormatHMS(Value, ":", ":", String.Empty)
        End Function

        '''<summary>Convert a value between 0 and 24 to 00:00:00 and 23:59:59.</summary>
        Public Shared Function FormatHMS(ByVal Value As Double, ByVal H As String, ByVal M As String, ByVal S As String) As String
            Dim Hours As Integer = CInt(Fix(Value))
            Value = (Value - Hours) * 60
            Dim Minutes As Integer = CInt(Fix(Value))
            Value = (Value - Minutes) * 60
            Dim Seconds As Integer = CInt(Fix(Value))
            Return Hours.ToString.Trim & H & Format(Minutes, "00").Trim & M & Format(Seconds, "00").Trim & S
        End Function

        Public Shared Function Format360Degree(ByVal Degree As Double, ByVal SecRounding As Integer) As String
      Dim Sign As Integer = Math.Sign(Degree) : Degree = Math.Abs(Degree)
      Dim Deg As Integer = CInt(Math.Floor(Degree)) : Degree = (Degree - Deg) * 60
      Dim Min As Integer = CInt(Math.Floor(Degree)) : Degree = (Degree - Min) * 60
      Dim Sec As Double = Degree
            Return CStr(IIf(Sign = -1, "-", "+")) & Deg.ToString.Trim & "° " & Min.ToString.Trim & "' " & Format(Math.Round(Sec, SecRounding), "#0.00").Replace(",", ".").Trim & """"
        End Function

    Public Shared Function DateTimeForCSV(ByVal Value As DateTime) As String
      Return Format(Value, "dd.MM.yyyy HH:mm:ss")
    End Function

    Public Shared Function DateTimeForDisplay(ByVal Value As DateTime) As String
      Return Format(Value, "yyyy.MM.ddTHH:mm:ss")
    End Function

    Public Shared Function TimeForDisplay(ByVal Value As DateTime) As String
      Return Format(Value, "HH:mm:ss")
    End Function

  End Class

End Namespace