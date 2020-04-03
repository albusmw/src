Option Explicit On
Option Strict On

Public Class cGenerics

    '''<summary>Sort the passed dictionary according to T1 (key).</summary>
    Public Shared Function GetDictionaryKeyElement(Of T1, T2)(ByRef Hist As Collections.Generic.Dictionary(Of T1, T2), ByVal Index As Integer) As T1
        Dim Keys(Hist.Keys.Count - 1) As T1 : Hist.Keys.CopyTo(Keys, 0)
        Return Keys(Index)
    End Function

End Class