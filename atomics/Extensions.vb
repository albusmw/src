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

    '''<summary>Returns the string the comes after the passed part.</summary>
    '''<param name="InputString"></param>
    '''<param name="Part"></param>
    '''<returns></returns>
    <Extension()>
    Public Function PartBefore(ByVal InputString As String, ByVal Part As String) As String
        Dim Pos As Integer = InputString.IndexOf(Part)
        If Pos > 0 Then
            Return InputString.Substring(0, Pos)
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
    Public Function ValRegIndep(ByVal Value As Double, ByVal Format As String) As String
        Return Value.ToString(Format, Globalization.CultureInfo.InvariantCulture).Trim
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

Module VectorExtension

    <Extension()>
    Public Function ToDouble(ByVal Vector As Collections.Generic.List(Of UInt32)) As Double()
        Dim RetVal(Vector.Count - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Vector.Item(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Vector() As UInt32) As Double()
        Dim RetVal(Vector.GetUpperBound(0) - 1) As Double
        Threading.Tasks.Parallel.For(0, Vector.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Vector(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Vector As List(Of Long)) As Double()
        Dim RetVal(Vector.Count - 1) As Double
        Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                         RetVal(Idx) = Vector.Item(Idx)
                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Vector() As Long) As Double()
        Dim RetVal(Vector.GetUpperBound(0) - 1) As Double
        Parallel.For(0, Vector.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                         RetVal(Idx) = Vector(Idx)
                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
        Dim RetVal(Vector.Count - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Vector.Item(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Vector() As ULong) As Double()
        Dim RetVal(Vector.GetUpperBound(0) - 1) As Double
        Threading.Tasks.Parallel.For(0, Vector.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Vector(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

End Module
