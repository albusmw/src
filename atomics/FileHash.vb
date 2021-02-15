Option Explicit On
Option Strict On

Public Class FileHash

    '''<summary>Calculate the MD5 checksum of the given file.</summary>
    '''<param name="FileName">File to compute hash for.</param>
    '''<returns>Hash with - in between.</returns>
    Public Shared Function MD5(ByVal FileName As String) As String
        If System.IO.File.Exists(FileName) = False Then Return String.Empty
        Using CheckSumCalc As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
            Using stream = System.IO.File.OpenRead(FileName)
                Return BitConverter.ToString(CheckSumCalc.ComputeHash(stream))
            End Using
        End Using
    End Function

    '''<summary>Calculate the MD5 checksum of the given file.</summary>
    '''<param name="FileName">File to compute hash for.</param>
    '''<returns>Hash with - in between.</returns>
    Public Shared Function SHA256(ByVal FileName As String) As String
        Using CheckSumCalc As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            Using stream = System.IO.File.OpenRead(FileName)
                Return BitConverter.ToString(CheckSumCalc.ComputeHash(stream))
            End Using
        End Using
    End Function

End Class