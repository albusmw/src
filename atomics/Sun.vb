Option Explicit On
Option Strict On

Namespace AstroCalc.NET

  Public Class Sun

    Private Const DegToRad As Double = System.Math.PI / 180
    Private Const RadToDeg As Double = 180 / System.Math.PI

        Public Structure sSunRaiseAndSet
            Public SunRaise As DateTime
            Public SunSet As DateTime
            Public RaiseCivil As DateTime
            Public RaiseNautical As DateTime
            Public RaiseAstronomical As DateTime
            Public DawnCivil As DateTime
            Public DawnNautical As DateTime
            Public DawnAstronomical As DateTime
        End Structure

        '''<summary>Return TRUE is the sun is visible, FALSE else.</summary>
        Public Shared Function SunVisible(ByVal Longitude As Double, ByVal Latitude As Double, ByVal Moment As DateTime) As Boolean
            Dim Dummy As Double
            Dim Sol_Height As Double = Double.NaN
            AstroCalc.NET.Sun.SunPos(Moment, Longitude, Latitude, Dummy, Sol_Height)
            If Sol_Height < 0 Then Return False Else Return True
        End Function

        '''<summary>Calculate the parameters for the (next) night.</summary>
        Public Shared Function NightPreCalc(ByVal Longitude As Double, ByVal Latitude As Double) As sSunRaiseAndSet
            Dim RetVal As sSunRaiseAndSet
            Dim Param1 As sSunRaiseAndSet
            Dim Param2 As sSunRaiseAndSet
            If SunVisible(Longitude, Latitude, Now) = False Then
                'We are in the night
                If Now.Hour > 12 Then
                    'Sun raise is the next day
                    Param1 = SunRaiseAndSet(Longitude, Latitude, Now)
                    Param2 = SunRaiseAndSet(Longitude, Latitude, Now.AddDays(1))
                Else
                    'Sun raise is the this day
                    Param1 = SunRaiseAndSet(Longitude, Latitude, Now.AddDays(-1))
                    Param2 = SunRaiseAndSet(Longitude, Latitude, Now)
                End If
            Else
                'Day time ...
                Param1 = SunRaiseAndSet(Longitude, Latitude, Now)
                Param2 = SunRaiseAndSet(Longitude, Latitude, Now.AddDays(1))
            End If
            RetVal.DawnAstronomical = Param1.DawnAstronomical
            RetVal.DawnCivil = Param1.DawnCivil
            RetVal.DawnNautical = Param1.DawnNautical
            RetVal.RaiseAstronomical = Param2.RaiseAstronomical
            RetVal.RaiseCivil = Param2.RaiseCivil
            RetVal.RaiseNautical = Param2.RaiseNautical
            RetVal.SunRaise = Param2.SunRaise
            RetVal.SunSet = Param1.SunSet
            Return RetVal
        End Function

        Public Shared Function SunRaiseAndSet(ByVal Longitude As Double, ByVal Latitude As Double) As sSunRaiseAndSet
            Return SunRaiseAndSet(Longitude, Latitude, Now)
        End Function

        Public Shared Function SunRaiseAndSet(ByVal Longitude As Double, ByVal Latitude As Double, ByVal StartDate As Date) As sSunRaiseAndSet

            Dim RetVal As sSunRaiseAndSet
            Dim LastHeight As Double

            Dim Year As Integer = StartDate.Year
            Dim Month As Integer = StartDate.Month
            Dim Day As Integer = StartDate.Day
            Dim Ticks As New TimeSpan(0, 0, 1)

            Dim Sol_Height As Double
            Dim Dummy As Double

            'Sunrise
            Dim SunRaise As DateTime = New DateTime(Year, Month, Day, 12, 0, 0)
            LastHeight = Double.NaN
            Do
                AstroCalc.NET.Sun.SunPos(SunRaise, Longitude, Latitude, Dummy, Sol_Height)
                If Double.IsNaN(LastHeight) = True Then
                    LastHeight = Sol_Height
                Else
                    If LastHeight > 0 And Sol_Height <= 0 Then RetVal.SunRaise = SunRaise
                    If LastHeight > -6 And Sol_Height <= -6 Then RetVal.RaiseCivil = SunRaise
                    If LastHeight > -12 And Sol_Height <= -12 Then RetVal.RaiseNautical = SunRaise
                    If LastHeight > -18 And Sol_Height <= -18 Then RetVal.RaiseAstronomical = SunRaise
                    LastHeight = Sol_Height
                End If
                SunRaise = SunRaise.Subtract(Ticks)
            Loop Until Sol_Height <= -18

            'Sunset
            Dim SunSet As DateTime = New DateTime(Year, Month, Day, 12, 0, 0)
            LastHeight = Double.NaN
            Do
                AstroCalc.NET.Sun.SunPos(SunSet, Longitude, Latitude, Dummy, Sol_Height)
                If Double.IsNaN(LastHeight) = True Then
                    LastHeight = Sol_Height
                Else
                    If LastHeight > 0 And Sol_Height <= 0 Then RetVal.SunSet = SunSet
                    If LastHeight > -6 And Sol_Height <= -6 Then RetVal.DawnCivil = SunSet
                    If LastHeight > -12 And Sol_Height <= -12 Then RetVal.DawnNautical = SunSet
                    If LastHeight > -18 And Sol_Height <= -18 Then RetVal.DawnAstronomical = SunSet
                    LastHeight = Sol_Height
                End If
                SunSet = SunSet.Add(Ticks)
            Loop Until Sol_Height <= -18

            Return RetVal

        End Function

        Public Shared Sub SunPos(ByVal DT As DateTime, ByVal Longitude As Double, ByVal Latitude As Double, ByRef Azimut As Double, ByRef Height As Double)

      'Ensure UTC
      DT = DT.ToUniversalTime

      'Ekliptikalkoordinate der Sonne 
      Dim n As Double = DateAndTime.JDUT(DT) - 2451545                            'Tage seit dem Standardäquinoktium J2000.0
      Dim L As Double = Degree(280.45999999999998 + (0.98564739999999995 * n))    'mittlere ekliptikale Länge L
      Dim g As Double = Degree(357.52800000000002 + (0.98560029999999998 * n))    'mittlere Anomalie g
      Dim A As Double = L + (1.915 * Sin(g)) + (0.02 * Sin(2 * g))                'ekliptikale Länge A

      'Äquatorialkoordinaten der Sonne 
      Dim e As Double = 23.439 - 0.00000039999999999999998 * n                    'Schiefe der Ekliptik e
      Dim Rekta As Double = ArcTan(Cos(e) * Sin(A) / Cos(A))                      'Rektaszension alpha
      If Cos(A) < 0 Then Rekta += 180
      Dim Dekli As Double = ArcSin(Sin(e) * Sin(A))                               'Deklination delta

      'Horizontalkoordinaten der Sonne 
      Dim JD0 As Double = DateAndTime.JDUT(DT.Year, DT.Month, DT.Day, 0, 0, 0)    'Julianische Tageszahl
      Dim T0 As Double = (JD0 - 2451545) / 36525                                  'Julianischen Jahrhunderte ab J2000.0
      Dim T As Double = DT.Hour + (DT.Minute / 60) + (DT.Second / 3600)
      Dim StarTime As Double = 6.6973760000000002 + (2400.05134 * T0) + (1.0027379999999999 * T)  'mittlere Sternzeit in Greenwich
      StarTime = Hours(StarTime)
      Dim HourAngle As Double = StarTime * 15                                     'Greenwich-Stundenwinkel
      Dim Spring As Double = HourAngle + Longitude                                'Stundenwinkel des Frühlingspunkts
      Dim SolAngle As Double = Spring - Rekta                                     'Stundenwinkel der Sonne

      Dim Azimut_down As Double = ((Cos(SolAngle) * Sin(Latitude)) - (Tan(Dekli) * Cos(Latitude)))
      Azimut = ArcTan(Sin(SolAngle) / Azimut_down)
      If Azimut_down < 0 Then Azimut += 180
      Height = ArcSin((Cos(Dekli) * Cos(SolAngle) * Cos(Latitude)) + (Sin(Dekli) * Sin(Latitude)))

    End Sub

    '''<summary>Ensure the value to be within 0..360 degree.</summary>
    '''<param name="X">X [degree].</param>
    '''<returns>X be within 0..360 degree.</returns>
    Private Shared Function Degree(ByVal X As Double) As Double
      If X > 360 Then
        Do
          X -= 360
        Loop Until X <= 360
      Else
        If X < 0 Then
          Do
            X += 360
          Loop Until X >= 0
        End If
      End If
      Return X
    End Function

    '''<summary>Ensure the value to be within 0..24 hours.</summary>
    '''<param name="X">X [hours].</param>
    '''<returns>X be within 0..24 hours.</returns>
    Private Shared Function Hours(ByVal X As Double) As Double
      If X > 24 Then
        Do
          X -= 24
        Loop Until X <= 24
      Else
        If X < 0 Then
          Do
            X += 24
          Loop Until X >= 0
        End If
      End If
      Return X
    End Function

    '''<summary>Calculate sine [degree].</summary>
    '''<param name="X_deg">X [degree].</param>
    '''<returns>Sin(x).</returns>
    Private Shared Function Sin(ByVal X_deg As Double) As Double
      Return Math.Sin(X_deg * DegToRad)
    End Function

    '''<summary>Calculate cosine [degree].</summary>
    '''<param name="X_deg">X [degree].</param>
    '''<returns>Cos(x).</returns>
    Private Shared Function Cos(ByVal X_deg As Double) As Double
      Return Math.Cos(X_deg * DegToRad)
    End Function

    '''<summary>Calculate tangens [degree].</summary>
    '''<param name="X">X [degree].</param>
    '''<returns>Tan(x).</returns>
    Private Shared Function Tan(ByVal X As Double) As Double
      Return Math.Tan(X * DegToRad)
    End Function

    '''<summary>Calculate arcus tangens [degree].</summary>
    '''<param name="X_rad">X [rad].</param>
    '''<returns>Cos(x).</returns>
    Private Shared Function ArcTan(ByVal X_rad As Double) As Double
      Return RadToDeg * Math.Atan(X_rad)
    End Function

    '''<summary>Calculate arcus sinus [degree].</summary>
    '''<param name="X_rad">X [rad].</param>
    '''<returns>Cos(x).</returns>
    Private Shared Function ArcSin(ByVal X_rad As Double) As Double
      Return RadToDeg * Math.Asin(X_rad)
    End Function

    Private Class DateAndTime

      '''<summary>Calculate the Julian Date from the given date and time in Universal Time.</summary>
      '''<param name="Value">Date value in Universal Time.</param>
      '''<returns>Julian Date.</returns>
      '''<remarks>See http://de.wikipedia.org/wiki/Julianisches_Datum#Astronomisches_Julianisches_Datum for details.</remarks>
      Public Shared Function JDUT(ByVal Value As DateTime) As Double
        Return JDUT(Value.Year, Value.Month, Value.Day, Value.Hour, Value.Minute, Value.Second)
      End Function

      '''<summary>Calculate the Julian Date from the given date and time in Universal Time.</summary>
      '''<param name="Year">Year.</param>
      '''<param name="Month">Month.</param>
      '''<param name="Day">Day.</param>
      '''<param name="Hour">Hour.</param>
      '''<param name="Minute">Minute.</param>
      '''<param name="Second">Second.</param>
      '''<returns>Julian Date.</returns>
      '''<remarks>See http://de.wikipedia.org/wiki/Julianisches_Datum#Astronomisches_Julianisches_Datum for details.</remarks>
      Public Shared Function JDUT(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer) As Double

        Dim Y As Integer = Year
        Dim M As Integer = Month
        Dim D As Integer = Day

        Dim A As Integer = 0
        Dim B As Integer = 0

        Dim H As Double = (Hour / 24) + (Minute / 1440) + (Second / 86400)

        If M <= 2 Then
          Y -= 1 : M += 12
        End If

        If IsGregorian(Y, M, D) Then
          A = CInt(System.Math.Floor(Y / 100))
          B = CInt(2 - A + System.Math.Floor(A / 4))
        End If

        Return System.Math.Floor(365.25 * (Y + 4716)) + System.Math.Floor(30.600100000000001 * (M + 1)) + D + H + B - 1524.5

      End Function

      '''<summary>Determine if the date is in the gregorian calender.</summary>
      '''<returns>TRUE for gregorian, FALSE for julian. If neighter nor, an exception is thrown.</returns>
      Private Shared Function IsGregorian(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer) As Boolean
        Select Case Year
          Case Is > 1582 : Return True
          Case Is < 1582 : Return False
          Case Else
            Select Case Month
              Case Is > 10 : Return True
              Case Is < 10 : Return False
              Case Else
                Select Case Day
                  Case Is >= 15 : Return True
                  Case Is <= 4 : Return False
                  Case Else : Throw New Exception("Date is not a valid date in gregorian or julian calender")
                End Select
            End Select
        End Select

      End Function

    End Class

  End Class

End Namespace
