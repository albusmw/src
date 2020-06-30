Option Explicit On
Option Strict On

Partial Public Class Parser

    Public Shared Event [Error](ByVal Text As String)

    '''<summary>Different fields in the SQM log.</summary>
    '''<see cref="https://knightware.biz/sqm/SqmPro3.pdf"/>
    Public Class cSQMFields
        Public Shared ReadOnly Property Invalid As Integer = Integer.MinValue
        ''' <summary>Year/Month/Day, is PC (local) time when the reading was acquired.</summary>
        Public [Date] As Integer = 0
        ''' <summary>Hour/Minute/Second is PC (local) time when the reading was acquired.</summary>
        Public [Time] As Integer = 1
        ''' <summary>Magnitudes per square arc second as acquired from the meter.</summary>
        Public [MPSAS] As Integer = Invalid
        ''' <summary>Calculated from the MPSAS value acquired from the meter.</summary>
        Public [NELM] As Integer = Invalid
        Public SolarAlt As Integer = Invalid
        Public LunarAlt As Integer = Invalid
        Public LunarPhase As Integer = Invalid
    End Class

    Public Structure sSQMToday
        Public MPSAS As Double
        Public SolAlt As Double
        Public MoonAlt As Double
        Public MoonPhase As Double
        Public Sub New(ByVal NewMPSAS As Double, ByVal NewSolAlt As Double, ByVal NewMoonAlt As Double, ByVal NewMoonPhase As Double)
            MPSAS = NewMPSAS
            SolAlt = NewSolAlt
            MoonAlt = NewMoonAlt
            MoonPhase = NewMoonPhase
        End Sub
    End Structure

    '''<summary>Parse the passed SQM file.</summary>
    '''<param name="Filename">Filename to parse.</param>
    '''<param name="TodayDate">Today's date to run calculation for.</param>
    '''<param name="BestValues">Dictioary of best values.</param>
    '''<param name="Today"></param>
    '''<returns>List of raw values only date-time-MPSAS</returns>
    Public Shared Function ParseSQM(ByVal Filename As String, ByVal TodayDate As Date, ByRef BestValues As Dictionary(Of Date, Ato.cSingleValueStatistics), ByRef Today As Dictionary(Of DateTime, sSQMToday)) As List(Of String)

        Dim Content As String() = Nothing
        Dim SplitChar As Char = ","c
        Dim RetVal As New List(Of String)

        'Try to load the complete content
        Try
            If System.IO.File.Exists(Filename) = False Then
                RaiseEvent [Error]("File <" & Filename & "> does not exist.")
                Return RetVal
            Else
                Content = System.IO.File.ReadAllLines(Filename)
            End If
        Catch ex As Exception
            RaiseEvent [Error]("File <" & Filename & "> could not be loaded.")
            Return RetVal
        End Try

        If IsNothing(Content) = True Then Return RetVal
        If Content.Length <= 2 Then Return RetVal

        Dim SQMFields As New cSQMFields
        Dim Headers As String() = Split(Content(1), SplitChar)
        For Idx As Integer = 0 To Headers.GetUpperBound(0)
            Select Case Headers(Idx).ToUpper
                Case "MPSAS".ToUpper : SQMFields.MPSAS = Idx
                Case "SolarAlt(deg)".ToUpper : SQMFields.SolarAlt = Idx
                Case "LunarAlt(deg)".ToUpper : SQMFields.LunarAlt = Idx
                Case "LunarPhase".ToUpper : SQMFields.LunarPhase = Idx
            End Select
        Next Idx

        For LineIdx As Integer = 0 To Content.GetUpperBound(0)
            If Content(LineIdx).StartsWith("20") Then
                Dim Line As String() = Split(Content(LineIdx), SplitChar)
                Dim DateTime As DateTime
                DateTime.TryParse(Line(SQMFields.Date) & " " & Line(SQMFields.Time), DateTime)
                Dim DateOnly As Date = DateTime.Date
                Dim MPSAS As Double = Val(Line(SQMFields.MPSAS).Replace(",", "."))
                RetVal.Add(DateTime.ValRegIndep & "__" & MPSAS.ValRegIndep)
                Dim SolAlt As Double = cSQMFields.Invalid : If SQMFields.SolarAlt <> cSQMFields.Invalid Then SolAlt = Val(Line(SQMFields.SolarAlt).Replace(",", "."))
                Dim LunarAlt As Double = cSQMFields.Invalid : If SQMFields.LunarAlt <> cSQMFields.Invalid Then LunarAlt = Val(Line(SQMFields.LunarAlt).Replace(",", "."))
                Dim LunarPhase As Double = cSQMFields.Invalid : If SQMFields.LunarPhase <> cSQMFields.Invalid Then LunarPhase = Val(Line(SQMFields.LunarPhase).Replace(",", "."))
                If SolAlt <= -6 Or SolAlt = cSQMFields.Invalid Then
                    If BestValues.ContainsKey(DateOnly) = False Then
                        BestValues.Add(DateOnly, New Ato.cSingleValueStatistics(Ato.cSingleValueStatistics.eValueType.Linear))
                    End If
                    BestValues(DateOnly).AddValue(MPSAS)
                End If
                'Display values 1 day before and after the specified one
                If Math.Abs(DateOnly.Subtract(TodayDate).TotalDays) <= 2 Then
                    Today.Add(DateTime, New sSQMToday(MPSAS, SolAlt, LunarAlt, LunarPhase))
                End If
            End If
        Next LineIdx

        Return RetVal

    End Function

End Class