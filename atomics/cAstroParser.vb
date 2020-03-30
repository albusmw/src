﻿Option Explicit On
Option Strict On

'''<summary>Class to parse right accension, decination, position data.</summary>
Public Class AstroParser

    '''<summary>Separator sign.</summary>
    Private Shared ReadOnly Sep As Char = CChar("|")

    '''<summary>NEW parser for coordinate input (not yet finished).</summary>
    Public Shared Function ParseCoord(ByVal Text As String) As Double

        Dim Opt As Text.RegularExpressions.RegexOptions = System.Text.RegularExpressions.RegexOptions.IgnoreCase
        Dim NumberRegEx As String = "[0-9.,]{1,}"
        Dim NonNumberRegEx As String = "[^0-9.,]{1,}"
        Dim Argument As String = "'22.7° 15m33.6''"
        Dim RegExInp As String = System.Text.RegularExpressions.Regex.Replace(Argument, "\s", String.Empty)
        Dim AllDigits As New Text.RegularExpressions.Regex(NumberRegEx & NonNumberRegEx, Opt)
        Dim Result As Text.RegularExpressions.MatchCollection = AllDigits.Matches(RegExInp)

        For Each Item As Text.RegularExpressions.Group In Result
            Dim NumPart As String = System.Text.RegularExpressions.Regex.Match(Item.Value, NumberRegEx).Value
            Dim NonNumPart As String = System.Text.RegularExpressions.Regex.Match(Item.Value, NonNumberRegEx).Value
            Console.WriteLine("<" & Item.Value & "> == <" & NumPart & "|" & NonNumPart & ">")
        Next Item

        Return Double.NaN

    End Function

    '''<summary>Returns the hours value from the given right accension value.</summary>
    '''<param name="Text">Text.</param>
    '''<returns>Value [h].</returns>
    Public Shared Function ParseRA(ByVal Text As String) As Double
        '1.) Try to generate a common notation
        Text = Text.Trim.ToUpper.Replace(" ", Sep).Replace(",", ".")
        Text = Text.Replace("H", Sep).Replace("'", Sep).Replace("""", Sep).Replace(":", Sep)
        Text = Text.Replace("′", Sep).Replace("″", Sep).Replace("D", Sep).Replace("MIN", Sep)
        Text = Text.TrimEnd(Sep)
        '2.) Calculate return value
        Dim Values As String() = Split(Text.TrimEnd(Sep), Sep)
        Select Case Values.Length
            Case 1
                Return (Val(Values(0)))
            Case 2
                Return (Val(Values(0)) + (Val(Values(1)) / 60))
            Case 3
                Return (Val(Values(0)) + (Val(Values(1)) / 60) + (Val(Values(2)) / 3600))
        End Select
        '4.) Conversion failed ...
        Return Double.NaN
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