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
    Public Function ToDouble(ByVal Argument As List(Of UInteger)) As Double()
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
    Public Function ToDouble(ByVal Argument As List(Of Long)) As Double()
        Dim RetVal(Argument.Count - 1) As Double
        Threading.Tasks.Parallel.For(0, RetVal.GetUpperBound(0) + 1, Sub(Idx As Integer)
                                                                         RetVal(Idx) = Argument(Idx)
                                                                     End Sub)

        Return RetVal
    End Function

End Module

Module DirectoryExtension

    '''<summary>Add the given value to the element or create a new element with this value.</summary>
    <Extension()>
    Public Sub AddTo(Of T1)(ByRef Dict As Dictionary(Of T1, UInt64), ByVal NewKey As T1, ByVal NewValue As UInt64)
        If IsNothing(Dict) = False Then
            If Dict.ContainsKey(NewKey) = False Then
                Dict.Add(NewKey, NewValue)
            Else
                Dict(NewKey) = Dict(NewKey) + NewValue
            End If
        End If
    End Sub

    '''<summary>Add the given value to the element or create a new element with this value.</summary>
    <Extension()>
    Public Sub AddTo(Of T1)(ByRef Dict As Dictionary(Of T1, UInt32), ByVal NewKey As T1, ByVal NewValue As UInt32)
        If IsNothing(Dict) = False Then
            If Dict.ContainsKey(NewKey) = False Then
                Dict.Add(NewKey, NewValue)
            Else
                Dict(NewKey) = Dict(NewKey) + NewValue
            End If
        End If
    End Sub

    '''<summary>Get a list of all keys in the dictionary passed.</summary>
    <Extension()>
    Public Function KeyList(Of T1, T2)(ByRef Dict As Dictionary(Of T1, T2)) As List(Of T1)
        If IsNothing(Dict) = True Then Return Nothing
        Return New List(Of T1)(Dict.Keys)
    End Function

    '''<summary>Get a list of all values in the dictionary passed.</summary>
    <Extension()>
    Public Function ValueList(Of T1, T2)(ByRef Dict As Dictionary(Of T1, T2)) As List(Of T2)
        Return New List(Of T2)(Dict.Values)
    End Function

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2)) As Dictionary(Of T1, T2)
        Return SortDictionary(Hist, False)
    End Function

    '''<summary>Sort the dictionary passed and invert the order.</summary>
    <Extension()>
    Public Function SortDictionaryInverse(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2)) As Dictionary(Of T1, T2)
        Return SortDictionary(Hist, True)
    End Function

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Private Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2), ByVal Reverse As Boolean) As Dictionary(Of T1, T2)

        'Generate a list
        Dim KeyList As New List(Of T1)
        For Each Entry As T1 In Hist.Keys
            KeyList.Add(Entry)
        Next Entry
        'Sort keys
        KeyList.Sort()
        If Reverse Then KeyList.Reverse()
        'Re-generate dictionary
        Dim RetVal As New Dictionary(Of T1, T2)
        For Each Entry As T1 In KeyList
            RetVal.Add(Entry, Hist(Entry))
        Next Entry
        Return RetVal

    End Function

End Module

Module ListExtensions

    '''<summary>Add the passed new element if it does not already exist.</summary>
    <Extension()>
    Public Sub AddNew(Of T1)(ByRef L As List(Of T1), ByVal NewElement As T1)
        If IsNothing(L) = False Then
            If L.Contains(NewElement) = False Then L.Add(NewElement)
        End If
    End Sub

End Module