Option Explicit On
Option Strict On

'''<summary>Get parts of the catalog strings (as described in the format description) - avoids mis-calculation, ....</summary>
Public Class CatParser

  '''<summary>Return value of an invalid int conversion.</summary>
  Public Const IntInvalid As Integer = -1
  '''<summary>Return value of an invalid double conversion.</summary>
  Public Const DblInvalid As Double = Double.NaN

  '''<summary>Process the information "34-36" from the format descriptions.</summary>
  '''<param name="Text">Text to get information from.</param>
  '''<param name="Position">Starting position (1-based).</param>
  '''<returns>String value.</returns>
  Public Shared Function GetString(ByRef Text As String, ByVal Position As Integer) As String
    Dim CuttedText As String = Text.Substring(Position, 1).Trim
    Return CuttedText
  End Function

  '''<summary>Process the information "34-36" from the format descriptions.</summary>
  '''<param name="Text">Text to get information from.</param>
  '''<param name="From">Starting position (1-based).</param>
  '''<param name="To">End position (1-based).</param>
  '''<returns>String value.</returns>
  Public Shared Function GetString(ByRef Text As String, ByVal [From] As Integer, ByVal [To] As Integer) As String
    Dim CuttedText As String = Text.Substring(From - 1, [To] - From + 1).Trim
    Return CuttedText
  End Function

  '''<summary>Process the information "34-36" from the format descriptions.</summary>
  '''<param name="Text">Text to get information from.</param>
  '''<param name="From">Starting position (1-based).</param>
  '''<param name="To">End position (1-based).</param>
  '''<returns>Integer value.</returns>
  Public Shared Function GetInt(ByRef Text As String, ByVal [From] As Integer, ByVal [To] As Integer) As Integer
    Dim CuttedText As String = Text.Substring(From - 1, [To] - From + 1).Trim
    If CuttedText.Length > 0 Then
      Return CInt(CuttedText)
    Else
      Return IntInvalid
    End If
  End Function

  '''<summary>Process the information "34-36" from the format descriptions.</summary>
  '''<param name="Text">Text to get information from.</param>
  '''<param name="From">Starting position (1-based).</param>
  '''<param name="To">End position (1-based).</param>
  '''<returns>Double value.</returns>
  Public Shared Function GetFloat(ByRef Text As String, ByVal [From] As Integer, ByVal [To] As Integer) As Double
    Dim CuttedText As String = Text.Substring(From - 1, [To] - From + 1).Trim
    If CuttedText.Length > 0 Then
      Return CDbl(CuttedText)
    Else
      Return DblInvalid
    End If
  End Function

End Class
