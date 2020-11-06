Option Explicit On
Option Strict On

'''<summary>Class to parse right accension, decination, position data, ....</summary>
Public Class AstroParser

    '''<summary>Result of the RegEx position parsing.</summary>
    Public Structure sCoordParts
        Public Value As Double
        Public Unit As String
        Public Sub New(ByVal NewValue As Double, ByVal NewUnit As String)
            Value = NewValue
            Unit = NewUnit
        End Sub
    End Structure

    '''<summary>Separator sign.</summary>
    Private Shared ReadOnly Sep As Char = CChar("|")

    '''<summary>NEW parser for coordinate input (not yet finished) - split entry into number and unit parts.</summary>
    '''<param name="Argument">Text to parse.</param>
    '''<param name="NoUnits">Indicate that there are no units at all.</param>
    '''<returns>List of found digit-unit pairs.</returns>
    Public Shared Function ParseCoord(ByVal Argument As String, ByVal NoUnits As Boolean) As List(Of sCoordParts)

        Dim Opt As Text.RegularExpressions.RegexOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase
        Dim NumberRegEx As String = "[0-9.,]{1,}"
        Dim NonNumberRegEx As String = "[^0-9.,]{0,}"
        Dim RegExInp As String = System.Text.RegularExpressions.Regex.Replace(Argument, "\s", String.Empty)
        Dim AllDigits As New Text.RegularExpressions.Regex(NumberRegEx & NonNumberRegEx, Opt)
        Dim Result As Text.RegularExpressions.MatchCollection = AllDigits.Matches(RegExInp)

        Dim RetVal As New List(Of sCoordParts)
        NoUnits = True
        For Each Item As Text.RegularExpressions.Group In Result
            Dim NumPart As String = System.Text.RegularExpressions.Regex.Match(Item.Value, NumberRegEx).Value
            Dim NonNumPart As String = System.Text.RegularExpressions.Regex.Match(Item.Value, NonNumberRegEx).Value.Trim
            RetVal.Add(New sCoordParts(Val(NumPart.Replace(",", ".")), NonNumPart))
            If NonNumPart.Length > 0 Then NoUnits = False
        Next Item

        Return RetVal

    End Function

    '''<summary>Returns the hours value from the given right accension value.</summary>
    '''<param name="Text">Text.</param>
    '''<returns>Value [h].</returns>
    Public Shared Function ParseRA(ByVal Text As String) As Double
        'In case of no units passen (e.g. ":"), HH:MM:SS is assumed and default units are used
        Dim NoUnits As Boolean = True
        Dim DefaultUnits As String() = {"H", "M", "S"}
        Dim UnitPtr As Integer = 0

        Dim Splitted As List(Of sCoordParts) = ParseCoord(Text, NoUnits)
        Dim RetVal As Double = 0
        For Each Part As sCoordParts In Splitted
            Dim UnitToUse As String = Part.Unit.ToUpper
            If NoUnits Then UnitToUse = DefaultUnits(UnitPtr)
            Select Case UnitToUse
                Case "H" : RetVal += Part.Value
                Case "M", "MIN", "'", "´" : RetVal += (Part.Value / 60.0)
                Case "S", "SEC", "''", Chr(34), "´´", "``" : RetVal += (Part.Value / 3600.0)
            End Select
            UnitPtr += 1
        Next Part
        Return RetVal
    End Function

    '''<summary>Returns the degree value from the given text.</summary>
    '''<param name="Text">Text.</param>
    '''<returns>Value [°].</returns>
    Public Shared Function ParseDeclination(ByVal Text As String) As Double
        '1.) Try to generate a common notation
        Text = Text.ToUpper.Replace(" ", String.Empty).Replace(",", ".").Replace(":", Sep)
        Text = Text.Replace("°", Sep).Replace("'", Sep).Replace("""", Sep)                      'common notation
        Text = Text.Replace("′", Sep).Replace("″", Sep).Replace("D", Sep).Replace("MIN", Sep)   'special minus and second signs / texts
        Text = Text.Replace("−", "-").Replace("–", "-").Replace("+", String.Empty)              'special "minus" signals and plus sign
        Text = Text.TrimEnd(Sep)
        '2.) Decide negative
        Dim Sign As Double = 1
        If Text.StartsWith("-") Then
            Sign = -1 : Text = Text.Substring(1)
        End If
        '3.) Calculate return value
        Dim Values As String() = Split(Text, Sep)
        Select Case Values.Length
            Case 1
                Return Sign * (Val(Values(0)))
            Case 2
                Return Sign * ((Val(Values(0)) + (Val(Values(1)) / 60)))
            Case 3
                Return Sign * ((Val(Values(0)) + (Val(Values(1)) / 60) + (Val(Values(2)) / 3600)))
        End Select
        '3.) Conversion failed ...
        Return Double.NaN
    End Function

    '''<summary>Returns the degree value from the given latitude or longitude value.</summary>
    '''<param name="Text">Text.</param>
    '''<returns>Value [°].</returns>
    Public Shared Function ParsePosition(ByVal Text As String) As Double
        '1.) Try to generate a common notation
        Text = Text.Trim.ToUpper.Replace(" ", Sep).Replace(",", ".")
        Text = Text.Replace("°", Sep).Replace("'", Sep).Replace("""", Sep).Replace(":", Sep)
        Text = Text.Replace("′", Sep).Replace("″", Sep).Replace("D", Sep).Replace("MIN", Sep)
        Text = Text.TrimEnd(Sep)
        '2.) Decide north or south
        Dim Sign As Double = 1
        Text = Text.Replace("SOUTH", "S").Replace("NORTH", "N").Replace("WEST", "W").Replace("EAST", "E")
        If Text.EndsWith("S") Or Text.EndsWith("W") Then
            Sign = -1 : Text = Text.Substring(0, Text.Length - 1)
        End If
        If Text.EndsWith("N") Or Text.EndsWith("E") Then
            Sign = 1 : Text = Text.Substring(0, Text.Length - 1)
        End If
        If Text.StartsWith("-") Then
            Sign = -1 : Text = Text.Substring(1)
        End If
        '3.) Calculate return value
        Dim Values As String() = Split(Text.TrimEnd(Sep), Sep)
        Select Case Values.Length
            Case 1
                Return Sign * (Val(Values(0)))
            Case 2
                Return Sign * (Val(Values(0)) + (Val(Values(1)) / 60))
            Case 3
                Return Sign * (Val(Values(0)) + (Val(Values(1)) / 60) + (Val(Values(2)) / 3600))
        End Select
        '4.) Conversion failed ...
        Return Double.NaN
    End Function

End Class