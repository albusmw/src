Option Explicit On
Option Strict On

'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

    '''<summary>Scan directory recursive.</summary>
    Public Class RecursivDirScanner

        Public Event CurrentlyScanning(ByVal DirectoryName As String)
        Public Event NewFileFound(ByVal FileName As String)
        Public AllFiles As New List(Of String)

        Public Property MustContain() As String
            Get
                Return MyMustContain
            End Get
            Set(value As String)
                MyMustContain = value
            End Set
        End Property
        Private MyMustContain As String = String.Empty

        Public Property MustContainCaseSensitive() As Boolean
            Get
                Return MyMustContainCaseSensitive
            End Get
            Set(value As Boolean)
                MyMustContainCaseSensitive = value
            End Set
        End Property
        Private MyMustContainCaseSensitive As Boolean = False

        Public Sub New(ByVal RootDirectory As String)
            MyRootDirectory = RootDirectory
        End Sub
        Private MyRootDirectory As String = String.Empty

        Public Sub Scan(ByVal Filter As String)
            AllFiles.Clear()
            InnerScanner(MyRootDirectory, "*", Filter, Integer.MaxValue)
        End Sub

        Public Sub Scan(ByVal Filter As String, ByVal MaxDepth As Integer)
            AllFiles.Clear()
            InnerScanner(MyRootDirectory, "*", Filter, MaxDepth)
        End Sub

        Public Sub Scan(ByVal DirFilter As String, ByVal FileFilter As String)
            AllFiles.Clear()
            InnerScanner(MyRootDirectory, DirFilter, FileFilter, Integer.MaxValue)
        End Sub

        Public Sub Scan(ByVal Root As String, ByVal DirFilter As String, ByVal FileFilter As String)
            AllFiles.Clear()
            InnerScanner(Root, DirFilter, FileFilter, Integer.MaxValue)
        End Sub

        Private Sub InnerScanner(ByVal Root As String, ByVal DirFilter As String, ByVal FileFilter As String, ByVal MaxDepth As Integer)

            'Scan files
            For Each File As String In System.IO.Directory.GetFiles(Root, FileFilter)
                If ValidFile(File) = True Then NewFile(File)
            Next File

            'Scan directories
            If MaxDepth > 0 Then
                For Each Directory As String In System.IO.Directory.GetDirectories(Root, DirFilter)
                    InnerScanner(Directory, DirFilter, FileFilter, MaxDepth - 1)
                Next Directory
            End If

        End Sub

        Private Function ValidFile(ByVal FileName As String) As Boolean
            If String.IsNullOrEmpty(MustContain) = True Then Return True
            If Me.MustContainCaseSensitive = False Then
                If System.IO.File.ReadAllText(FileName).Contains(MustContain) = True Then Return True
            Else
                If System.IO.File.ReadAllText(FileName).ToUpperInvariant.Contains(MustContain.ToUpperInvariant) = True Then Return True
            End If
            Return False
        End Function

        Private Sub NewFile(ByVal FileName As String)
            AllFiles.Add(FileName)
            RaiseEvent NewFileFound(FileName)
        End Sub

    End Class

End Namespace