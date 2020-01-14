Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices

Module StringExtension

    '''<summary>Returns the string the comes after the passed part.</summary>
    '''<param name="InputString"></param>
    '''<param name="Part"></param>
    '''<returns></returns>
    <Extension()>
    Public Function PartAfter(ByVal InputString As String, ByVal Part As String) As String
        Dim Pos As Integer = InputString.IndexOf(Part)
        If Pos > 0 Then
            Return InputString.Substring(Pos + Part.Length)
        Else
            Return String.Empty
        End If
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Single) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Double) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Byte) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Int16) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As UInt16) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Int32) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As UInt32) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Int64) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As UInt64) As String
        Return Value.ToString(Globalization.CultureInfo.InvariantCulture).Trim
    End Function

End Module
