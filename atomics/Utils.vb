Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It is NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    Public Class Utils

        Public Shared Function CreateDesktopShortcut() As Boolean

            '''<summary>Create a shortcut to the current running EXE.</summary>
            Try
                Dim DesktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim EXEName As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                Dim EXENameOnly As String = System.IO.Path.GetFileNameWithoutExtension(EXEName)
                Dim Shell32Type As Type = Type.GetTypeFromProgID("WScript.Shell", True)
                Dim Shell32Instance As Object = Activator.CreateInstance(Shell32Type)
                Dim InvokeResult As Object = Shell32Type.InvokeMember("CreateShortcut", Reflection.BindingFlags.InvokeMethod, Nothing, Shell32Instance, New Object() {DesktopPath & "\" & EXENameOnly & ".lnk"})
                Dim X As Object = InvokeResult.GetType.InvokeMember("TargetPath", Reflection.BindingFlags.SetProperty, Nothing, InvokeResult, New Object() {EXEName})
                Dim Y As Object = InvokeResult.GetType.InvokeMember("WindowStyle", Reflection.BindingFlags.SetProperty, Nothing, InvokeResult, New Object() {1})
                Dim Z As Object = InvokeResult.GetType.InvokeMember("Save", Reflection.BindingFlags.InvokeMethod, Nothing, InvokeResult, Nothing)
                Return True
            Catch ex As Exception
                Return False
            End Try

        End Function

        Public Shared Function GetURLContent(ByVal RequestURL As String) As String
            'Query data from request URL
            Dim address As New Uri(RequestURL)
            Dim mWC As New System.Net.WebClient
            Dim InStream As IO.Stream = mWC.OpenRead(address)
            Dim InReader As New IO.StreamReader(InStream)
            Dim Data As String = InReader.ReadToEnd
            Return Data
        End Function

    End Class

End Namespace