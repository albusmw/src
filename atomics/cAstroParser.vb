Option Explicit On
Option Strict On

Public Class AstroParser

    '''<summary>Separator sign.</summary>
    Private Shared ReadOnly Sep As Char = CChar("|")

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