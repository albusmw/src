Option Explicit On
Option Strict On

Public Class cVizierQuery

    Public Shared Function GetNamedElement(ByVal ElementName As String, ByRef RA As Double, ByRef Dec As Double) As String

        RA = Double.NaN
        Dec = Double.NaN

        Dim RootURL As String = "http://simbad.u-strasbg.fr/simbad/sim-id"
        RootURL &= "?Ident=" & ElementName & "&output.format=ASCII"

        Try
            Dim req As Net.HttpWebRequest = CType(Net.WebRequest.Create(RootURL), Net.HttpWebRequest)
            req.Method = "GET"
            Dim Answer As String = (New IO.StreamReader(CType(req.GetResponse(), Net.HttpWebResponse).GetResponseStream)).ReadToEnd
            Return Answer
        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

End Class