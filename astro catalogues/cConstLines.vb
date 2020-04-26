Option Explicit On
Option Strict On

Namespace AstroCalc.NET

  Namespace Databases

    Public Class cConstLines

      '''<summary>Get the constallation line for the requested constellation.</summary>
      '''<param name="IAUName">IAU name.</param>
      '''<returns>List of HD numbers - -1 indicated to insert a "break" in the line.</returns>
      Public Function GetConstLine(ByVal IAUName As String) As List(Of Integer)

        Dim RetVal As New List(Of Integer)

        Select Case IAUName.ToUpper

          Case "UMa".ToUpper
            'Ursa Major 
            RetVal.Add(120315)
            RetVal.Add(116656)
            RetVal.Add(112185)
            RetVal.Add(106591)
            RetVal.Add(95689)
            RetVal.Add(95418)
            RetVal.Add(103287)
            RetVal.Add(106591)

          Case Else

            'Do nothing...

        End Select

        Return RetVal

      End Function

    End Class

  End Namespace

End Namespace