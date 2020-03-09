Option Explicit On
Option Strict On

'''<summary>Class to get certain defined header information from a FITS file.</summary>
Public Class cFITSHeaderParser

    '''<summary>List of all header elements added.</summary>
    Private AllCards As New List(Of cFITSHeaderParser.sHeaderElement)

    '''<summary>If an element that already exists is added again, the already present entry is updated if this property is TRUE.</summary>
    Public Property UpdateExistingElement As Boolean = True

    '''<summary>Elements of one single FITS header line.</summary>
    Public Structure sHeaderElement
        Public Keyword As String
        Public Value As String
        Public Comment As String
        Public Sub New(ByVal NewKeyword As String, ByVal NewValue As String)
            Me.Keyword = NewKeyword
            Me.Value = NewValue
            Me.Comment = String.Empty
        End Sub
        Public Sub New(ByVal NewKeyword As String, ByVal NewValue As String, ByVal NewComment As String)
            Me.Keyword = NewKeyword
            Me.Value = NewValue
            Me.Comment = NewComment
        End Sub
        Public Sub New(ByVal NewKeyword As String, ByVal NewValue As Double)
            Me.Keyword = NewKeyword
            Me.Value = NewValue.ValRegIndep.Replace(",", ".")
            Me.Comment = String.Empty
        End Sub
        Public Sub New(ByVal NewKeyword As String, ByVal NewValue As Double, ByVal NewComment As String)
            Me.Keyword = NewKeyword
            Me.Value = NewValue.ValRegIndep.Replace(",", ".")
            Me.Comment = NewComment
        End Sub
        Public Shared Function Sorter(ByVal X As sHeaderElement, ByVal Y As sHeaderElement) As Integer
            Return X.Keyword.CompareTo(Y.Keyword)
        End Function
    End Structure

    '''<summary>Common FITS header information.</summary>
    Private Class cMyProps
        Public BitPix As Integer = 0
        Public BZERO As Double = 0.0
        Public BSCALE As Double = 1.0
        Public NAXIS As Integer = -1
        Public Width As Integer = -1
        Public Height As Integer = -1
        Public ColorValues As Integer = 0
        Public BytesPerSample As Integer = -1
        Public DataStartIdx As Integer = -1
    End Class
    Private MyProps As New cMyProps

    '''<summary>Bit per pixel - negative values indicate floating-point.</summary>
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

    Public ReadOnly Property NAXIS() As Integer
        Get
            Return MyProps.NAXIS
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

    Public Sub New(ByVal CardsToAdd As List(Of cFITSHeaderParser.sHeaderElement))
        AllCards = New List(Of cFITSHeaderParser.sHeaderElement)
        'Move through all elements and get known elements and convert them
        For Each Card As sHeaderElement In CardsToAdd
            'Check and update found element is exists
            Dim FoundIdx As Integer = ElementIdx(Card.Keyword)
            If FoundIdx = -1 Then
                AllCards.Add(Card)
            Else
                If UpdateExistingElement = True Then AllCards(FoundIdx) = Card
            End If
            Select Case Card.Keyword.Trim
                Case "BITPIX" : MyProps.BitPix = CInt(Card.Value)
                Case "NAXIS" : MyProps.NAXIS = CInt(Card.Value)
                Case "NAXIS1" : MyProps.Width = CInt(Card.Value)
                Case "NAXIS2" : MyProps.Height = CInt(Card.Value)
                Case "NAXIS3" : MyProps.ColorValues = CInt(Card.Value)
                Case "BZERO" : MyProps.BZERO = Val(Card.Value.Replace(",", "."))
                Case "BSCALE" : MyProps.BSCALE = Val(Card.Value.Replace(",", "."))
            End Select
        Next Card
    End Sub

    '''<summary>Add the card to the already existing list of cards.</summary>
    Public Sub Add(ByVal CardToAdd As cFITSHeaderParser.sHeaderElement)
        'Check and update found element is exists
        Dim FoundIdx As Integer = ElementIdx(CardToAdd.Keyword)
        If FoundIdx = -1 Then
            AllCards.Add(CardToAdd)
        Else
            If UpdateExistingElement = True Then AllCards(FoundIdx) = CardToAdd
        End If
        Dim Keyword As String = CardToAdd.Keyword.Trim
        Select Case Keyword
            Case "BITPIX" : MyProps.BitPix = CInt(ElementValue(Keyword))
            Case "NAXIS1" : MyProps.Width = CInt(ElementValue(Keyword))
            Case "NAXIS2" : MyProps.Height = CInt(ElementValue(Keyword))
            Case "NAXIS3" : MyProps.ColorValues = CInt(ElementValue(Keyword))
            Case "BZERO" : MyProps.BZERO = Val(ElementValue(Keyword).Replace(",", "."))
            Case "BSCALE" : MyProps.BSCALE = Val(ElementValue(Keyword).Replace(",", "."))
        End Select
    End Sub

    '''<summary>Return the value of the given keyword if present.</summary>
    '''<returns>Index of the found element or -1 if element is not found.</returns>
    Public Function ElementValue(ByVal Keyword As String) As String
        Keyword = Keyword.Trim.ToUpper
        For Idx As Integer = 0 To AllCards.Count - 1
            If AllCards(Idx).Keyword.Trim.ToUpper = Keyword Then Return AllCards(Idx).Value
        Next Idx
        Return Nothing
    End Function

    '''<summary>Check if an element with the given keyword already exists in the list of elements.</summary>
    '''<returns>Index of the found element or -1 if element is not found.</returns>
    Private Function ElementIdx(ByVal Keyword As String) As Integer
        Keyword = Keyword.Trim.ToUpper
        For Idx As Integer = 0 To AllCards.Count - 1
            If AllCards(Idx).Keyword.Trim.ToUpper = Keyword Then Return Idx
        Next Idx
        Return -1
    End Function

    '''<summary>Get a keyword-value dictionary.</summary>
    '''<remarks>If an entry is found again, the latest present entry in the list will be returned.</remarks>
    Public Function GetListAsDictionary() As Dictionary(Of String, Object)
        Dim RetVal As New Dictionary(Of String, Object)
        For Each Entry As cFITSHeaderParser.sHeaderElement In AllCards
            If RetVal.ContainsKey(Entry.Keyword) = False Then
                RetVal.Add(Entry.Keyword, Entry.Value)          'entry is new -> add
            Else
                RetVal(Entry.Keyword) = Entry.Value             'entry already exists -> update
            End If
        Next Entry
        Return RetVal
    End Function

    '''<summary>Get a keyword-value dictionary.</summary>
    '''<remarks>If an entry is found again, the latest present entry in the list will be returned.</remarks>
    Public Shared Function GetListAsDictionary(ByRef CardsToProcess As List(Of cFITSHeaderParser.sHeaderElement)) As Dictionary(Of String, Object)
        Dim RetVal As New Dictionary(Of String, Object)
        For Each Card As cFITSHeaderParser.sHeaderElement In CardsToProcess
            If RetVal.ContainsKey(Card.Keyword) = False Then
                RetVal.Add(Card.Keyword, Card.Value)          'entry is new -> add
            Else
                RetVal(Card.Keyword) = Card.Value             'entry already exists -> update
            End If
        Next Card
        Return RetVal
    End Function

End Class