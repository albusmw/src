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

Module DoubleExtension


    <Extension()>
    Public Function ValRegIndep(ByVal Value As String) As Double
        Return Val(Value.Replace(",", "."))
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

    ''''<summary>Get a list of all keys in the dictionary passed.</summary>
    <Extension()>
    Public Function ToDouble(ByRef Dict As Dictionary(Of Long, ULong).KeyCollection) As Double()
        Dim RetVal(Dict.Count - 1) As Double
        For Idx As Integer = 0 To Dict.Count - 1
            RetVal(Idx) = Dict(Idx)
        Next Idx
        Return RetVal
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
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2), ByVal Reverse As Boolean) As Dictionary(Of T1, T2)
        Dim DontCare1 As T1
        Dim DontCare2 As T1
        Return SortDictionary(Hist, Reverse, DontCare1, DontCare2)
    End Function

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Dictionary(Of T1, T2), ByVal Reverse As Boolean, ByRef Min As T1, ByRef Max As T1) As Dictionary(Of T1, T2)

        'Generate a list
        Dim KeyList As New List(Of T1)
        For Each Entry As T1 In Hist.Keys
            KeyList.Add(Entry)
        Next Entry
        'Sort keys
        KeyList.Sort()
        If KeyList.Count > 0 Then
            Min = KeyList(0)
            Max = KeyList(KeyList.Count - 1)
        End If
        If Reverse Then KeyList.Reverse()
        'Re-generate dictionary
        Dim RetVal As New Dictionary(Of T1, T2)
        For Each Entry As T1 In KeyList
            RetVal.Add(Entry, Hist(Entry))
        Next Entry
        Return RetVal

    End Function

End Module

Module ConcurrentDictionaryExtensions

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Concurrent.ConcurrentDictionary(Of T1, T2)) As Concurrent.ConcurrentDictionary(Of T1, T2)
        Return SortDictionary(Hist, False)
    End Function

    '''<summary>Sort the dictionary passed and invert the order.</summary>
    <Extension()>
    Public Function SortDictionaryInverse(Of T1, T2)(ByRef Hist As Concurrent.ConcurrentDictionary(Of T1, T2)) As Concurrent.ConcurrentDictionary(Of T1, T2)
        Return SortDictionary(Hist, True)
    End Function

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Concurrent.ConcurrentDictionary(Of T1, T2), ByVal Reverse As Boolean) As Concurrent.ConcurrentDictionary(Of T1, T2)
        Dim DontCare1 As T1
        Dim DontCare2 As T1
        Return SortDictionary(Hist, Reverse, DontCare1, DontCare2)
    End Function

    '''<summary>Sort the dictionary passed.</summary>
    <Extension()>
    Public Function SortDictionary(Of T1, T2)(ByRef Hist As Concurrent.ConcurrentDictionary(Of T1, T2), ByVal Reverse As Boolean, ByRef Min As T1, ByRef Max As T1) As Concurrent.ConcurrentDictionary(Of T1, T2)

        'Generate a list
        Dim KeyList As New List(Of T1)
        For Each Entry As T1 In Hist.Keys
            KeyList.Add(Entry)
        Next Entry
        'Sort keys
        KeyList.Sort()
        If KeyList.Count > 0 Then
            Min = KeyList(0)
            Max = KeyList(KeyList.Count - 1)
        End If
        If Reverse Then KeyList.Reverse()
        'Re-generate dictionary
        Dim RetVal As New Concurrent.ConcurrentDictionary(Of T1, T2)
        For Each Entry As T1 In KeyList
            RetVal.TryAdd(Entry, Hist(Entry))
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

Module DataGridViewExtensions

    '''<summary>Create an ASCII table from the passed table.</summary>
    <Extension()>
    Public Function ASCIITable(ByRef dgv As Windows.Forms.DataGridView) As List(Of String)

        'Get a list of all lines
        Dim Lines As New List(Of List(Of String))
        Dim EntryWidth As New List(Of Integer)
        For LineIdx As Integer = 0 To dgv.RowCount - 1
            Lines.Add(New List(Of String))
            For ColIdx As Integer = 0 To dgv.ColumnCount - 1
                Dim Item As Object = dgv.Item(ColIdx, LineIdx).Value
                Dim ItemString As String = String.Empty
                If IsNothing(Item) = False Then ItemString = Item.ToString
                Lines(LineIdx).Add(ItemString)
                If LineIdx = 0 Then EntryWidth.Add(ItemString.Length)
                If ItemString.Length > EntryWidth(ColIdx) Then EntryWidth(ColIdx) = ItemString.Length
            Next ColIdx
        Next LineIdx

        'Build the ASCII table
        Dim RetVal As New List(Of String)
        For Each Line As List(Of String) In Lines
            Dim OneLine As String = "|"
            For EntryIdx As Integer = 0 To Line.Count - 1
                OneLine &= Line(EntryIdx).PadLeft(EntryWidth(EntryIdx)) & "|"
            Next EntryIdx
            RetVal.Add(OneLine)
        Next Line

        Return RetVal

    End Function

End Module

Module DateTimeExtension

    '''<summary>Format the passed date to form a folder or file name part.</summary>
    <Extension()>
    Public Function ForFileSystem(ByVal Value As DateTime) As String
        Return Format(Value, "yyyy_MM_dd_HH_mm_ss")
    End Function

    '''<summary>Format the passed date for logging purpose (no date, 1/100 seconds).</summary>
    <Extension()>
    Public Function ForLogging(ByVal Value As DateTime) As String
        Return Format(Value, "HH.mm.ss:fff")
    End Function

    '''<summary>Format the passed date as ISO9660 format.</summary>
    '''<see cref="https://wiki.osdev.org/ISO_9660#Date.2Ftime_format"/>
    <Extension()>
    Public Function AsISO9660(ByVal Value As Date) As String
        Return Format(Value, "yyyyMMddHHmmssff")
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As Date) As String
        Return Format(Value, "yyyy-MM-ddTHH:mm:ss")
    End Function

    <Extension()>
    Public Function ValRegIndep(ByVal Value As TimeSpan) As String
        Dim Total As Double = Value.TotalSeconds / 3600
        Dim Hours As Long = CLng(Math.Floor(Total))
        Total = (Total - Hours) * 60
        Dim Minutes As Long = CLng(Math.Floor(Total))
        Total = (Total - Minutes) * 60
        Dim Seconds As Long = CLng(Math.Floor(Total))
        Return Format(Hours, "0").Trim & ":" & Format(Minutes, "00").Trim & ":" & Format(Seconds, "00").Trim
    End Function

End Module