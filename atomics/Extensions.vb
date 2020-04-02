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
    Public Function ToDouble(ByVal Argument As UInt16()) As Double()
        Dim RetVal(Argument.Length - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Argument As UInt32()) As Double()
        Dim RetVal(Argument.Length - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Argument As UInt64()) As Double()
        Dim RetVal(Argument.Length - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Argument As Collections.Generic.List(Of UInteger)) As Double()
        Dim RetVal(Argument.Count - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Argument As Long()) As Double()
        Dim RetVal(Argument.Length - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    <Extension()>
    Public Function ToDouble(ByVal Argument As Collections.Generic.List(Of Long)) As Double()
        Dim RetVal(Argument.Count - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

    '''<summary>Get a list of all keys in the dictionary passed.</summary>
    <Extension()>
    Public Function KeyList(Of T1, T2)(ByRef Dict As Collections.Generic.Dictionary(Of T1, T2)) As Collections.Generic.List(Of T1)
        Return New Collections.Generic.List(Of T1)(Dict.Keys)
    End Function

    '''<summary>Get a list of all values in the dictionary passed.</summary>
    <Extension()>
    Public Function ValueList(Of T1, T2)(ByRef Dict As Collections.Generic.Dictionary(Of T1, T2)) As Collections.Generic.List(Of T2)
        Return New Collections.Generic.List(Of T2)(Dict.Values)
    End Function

End Module
