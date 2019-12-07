'################################################################################
' !!! IMPORTANT NOTE !!!
' It it NOT ALLOWED that a member of ATO depends on any other file !!!
'################################################################################

Namespace Ato

  Public Class INIRead

    ''' <summary>Get the value of the INI file entry specified.</summary>
    ''' <param name="FileName">File NAME.</param>
    Public Shared Function GetINIFromFile(ByVal FileName As String, ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
      If System.IO.File.Exists(FileName) = True Then
        Return GetINIFromContent(System.IO.File.ReadAllBytes(FileName), Section, KeyName, DefaultValue)
      Else
        Return DefaultValue
      End If
    End Function

    ''' <summary>Get the value of the INI file content specified.</summary>
    ''' <param name="FileContent">File CONTENT.</param>
    ''' <param name="Section"></param>
    ''' <param name="KeyName"></param>
    ''' <param name="DefaultValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetINIFromContent(ByVal FileContent As Byte(), ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
      Return GetINIFromContent(System.Text.ASCIIEncoding.ASCII.GetString(FileContent), Section, KeyName, DefaultValue)
    End Function

    ''' <summary>Get the value of the INI file entry specified.</summary>
    ''' <param name="Content">File CONTENT.</param>
    Public Shared Function GetINIFromContent(ByVal Content As String, ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
      Dim CurrentSection As String = String.Empty
            For Each Line As String In Split(Content, System.Environment.NewLine)
                If Line.Trim.StartsWith("[") And Line.Trim.EndsWith("]") And Line.Trim.Length > 2 Then
                    CurrentSection = Line.Trim : CurrentSection = CurrentSection.Substring(1, CurrentSection.Length - 2)
                Else
                    Line = Line.TrimStart
                    Dim EqualPos As Integer = Line.IndexOf("=")
                    If EqualPos > 0 Then
                        Dim CurrentKeyName = Line.Substring(0, EqualPos)
                        If CurrentSection = Section And CurrentKeyName = KeyName Then
                            If Line.Length > EqualPos Then
                                Return Line.Substring(EqualPos + 1)
                            End If
                        End If
                    End If
                End If
            Next Line
            Return DefaultValue
    End Function

    


  End Class

End Namespace