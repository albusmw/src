Option Explicit On
Option Strict On

Public Class Everything

    Const DLLName As String = "Everything64.dll"

    <Runtime.InteropServices.DllImport(DLLName, CharSet:=Runtime.InteropServices.CharSet.Unicode)>
    Public Shared Sub Everything_SetSearchW(lpSearchString As String)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetMatchPath(bEnable As Boolean)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetMatchCase(bEnable As Boolean)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetMatchWholeWord(bEnable As Boolean)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetRegex(bEnable As Boolean)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetMax(dwMax As Integer)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetOffset(dwOffset As Integer)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetReplyWindow(hWnd As IntPtr)
    End Sub
    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SetReplyID(nId As Integer)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetMatchPath() As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetMatchCase() As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetMatchWholeWord() As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetRegex() As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetMax() As UInt32
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetOffset() As UInt32
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetSearch() As String
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetLastError() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetReplyWindow() As IntPtr
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetReplyID() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_QueryW(bWait As Boolean) As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_IsQueryReply(message As Integer, wParam As IntPtr, lParam As IntPtr, nId As UInteger) As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_SortResultsByPath()
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetNumFileResults() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetNumFolderResults() As Integer
    End Function
    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetNumResults() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetTotFileResults() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetTotFolderResults() As Integer
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_GetTotResults() As Integer
    End Function
    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_IsVolumeResult(nIndex As Integer) As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_IsFolderResult(nIndex As Integer) As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Function Everything_IsFileResult(nIndex As Integer) As Boolean
    End Function

    <Runtime.InteropServices.DllImport(DLLName, CharSet:=Runtime.InteropServices.CharSet.Unicode)>
    Public Shared Sub Everything_GetResultFullPathNameW(nIndex As Integer, lpString As Text.StringBuilder, nMaxCount As Integer)
    End Sub

    <Runtime.InteropServices.DllImport(DLLName)>
    Public Shared Sub Everything_Reset()
    End Sub

    '''<summary>Run a Everything search as in the Everything GUI.</summary>
    '''<param name="SearchQuery">e.g. "C:\GIT" IPP*.dll"</param>
    '''<returns></returns>
    Public Shared Function GetSearchResult(ByVal SearchQuery As String) As List(Of String)
        Dim RetVal As New List(Of String)
        Dim Errors As New List(Of Integer)
        Everything.Everything_SetSearchW(SearchQuery) : Errors.Add(Everything.Everything_GetLastError)
        Everything.Everything_QueryW(True) : Errors.Add(Everything.Everything_GetLastError)
        Dim NumberOfFiles As Integer = Everything.Everything_GetNumResults() : Errors.Add(Everything.Everything_GetLastError)
        'Get all found files
        Dim bufsize As Integer = 260
        Dim buf As New System.Text.StringBuilder(bufsize)
        For Idx As Integer = 0 To NumberOfFiles - 1
            Everything.Everything_GetResultFullPathNameW(Idx, buf, bufsize)
            RetVal.Add(buf.ToString)
        Next Idx
        Return RetVal
    End Function

    Const MY_REPLY_ID As Integer = 0

End Class
