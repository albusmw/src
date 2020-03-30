Option Explicit On
Option Strict On

Public Class cGenerics

    '''<summary>Sort the passed dictionary according to T1 (key).</summary>
    Public Shared Function SortDictionary(Of T1, T2)(ByRef Hist As Collections.Generic.Dictionary(Of T1, T2)) As Collections.Generic.Dictionary(Of T1, T2)

        'Generate a list
        Dim KeyList As New Collections.Generic.List(Of T1)
        For Each Entry As T1 In Hist.Keys
            KeyList.Add(Entry)
        Next Entry
        'Sort keys
        KeyList.Sort()
        'Re-generate dictionary
        Dim RetVal As New Collections.Generic.Dictionary(Of T1, T2)
        For Each Entry As T1 In KeyList
            RetVal.Add(Entry, Hist(Entry))
        Next Entry
        Return RetVal

    End Function

    '''<summary>Sort the passed dictionary according to T1 (key).</summary>
    Public Shared Function GetDictionaryKeyElement(Of T1, T2)(ByRef Hist As Collections.Generic.Dictionary(Of T1, T2), ByVal Index As Integer) As T1
        Dim Keys(Hist.Keys.Count - 1) As T1 : Hist.Keys.CopyTo(Keys, 0)
        Return Keys(Index)
    End Function

End Class