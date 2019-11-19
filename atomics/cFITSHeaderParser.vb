Option Explicit On
Option Strict On

'''<summary>Class to get certain defined header information from a FITS file.</summary>
Public Class cFITSHeaderParser

    '''<summary>Elements of one single FITS header line.</summary>
    Public Structure sHeaderElement
        Public Element As String
        Public Value As String
        Public Comment As String
    End Structure

    '''<summary>Common FITS header information.</summary>
    Private Class cMyProps
        Public BitPix As Integer = 0
        Public BZERO As Double = 0.0
        Public BSCALE As Double = 1.0
        Public Width As Integer = -1
        Public Height As Integer = -1
        Public ColorValues As Integer = 0
        Public BytesPerSample As Integer = -1
        Public DataStartIdx As Integer = -1
    End Class
    Private MyProps As New cMyProps

    Public ReadOnly Property BitPix() As Integer
        Get
            Return MyProps.BitPix
        End Get
    End Property

    Public ReadOnly Property BZERO() As Double
        Get
            Return MyProps.BZERO
        End Get
    End Property

    Public ReadOnly Property BSCALE() As Double
        Get
            Return MyProps.BSCALE
        End Get
    End Property

    Public ReadOnly Property Width() As Integer
        Get
            Return MyProps.Width
        End Get
    End Property

    Public ReadOnly Property Height() As Integer
        Get
            Return MyProps.Height
        End Get
    End Property

    Public ReadOnly Property BytesPerSample() As Integer
        Get
            Return Math.Abs(MyProps.BitPix) \ 8
        End Get
    End Property

    Public Sub New(ByVal HeaderElements As List(Of sHeaderElement))
        'Move through all elements and get known elements and convert them
        For Each NewHeaderElement As sHeaderElement In HeaderElements
            Select Case NewHeaderElement.Element
                Case "BITPIX" : MyProps.BitPix = CInt(NewHeaderElement.Value)
                Case "NAXIS1" : MyProps.Width = CInt(NewHeaderElement.Value)
                Case "NAXIS2" : MyProps.Height = CInt(NewHeaderElement.Value)
                Case "NAXIS3" : MyProps.ColorValues = CInt(NewHeaderElement.Value)
                Case "BZERO" : MyProps.BZERO = Val(NewHeaderElement.Value.Replace(",", "."))
                Case "BSCALE" : MyProps.BSCALE = Val(NewHeaderElement.Value.Replace(",", "."))
            End Select
        Next NewHeaderElement
    End Sub

End Class