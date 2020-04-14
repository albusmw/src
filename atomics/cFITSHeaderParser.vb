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
        Public Keyword As eFITSKeywords
        Public Value As Object
        Public Comment As String
        Public Sub New(ByVal NewKeyword As eFITSKeywords, ByVal NewValue As Object)
            Me.Keyword = NewKeyword
            Me.Value = NewValue
            Me.Comment = String.Empty
        End Sub
        Public Sub New(ByVal NewKeyword As eFITSKeywords, ByVal NewValue As Object, ByVal NewComment As String)
            Me.Keyword = NewKeyword
            Me.Value = NewValue
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
        Public NAXIS3 As Integer = 0
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

    Public ReadOnly Property NAXIS3() As Integer
        Get
            Return MyProps.NAXIS3
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
            Dim FoundIdx As Integer = IdxOfKeyword(Card.Keyword)
            If FoundIdx = -1 Then
                AllCards.Add(Card)
            Else
                If UpdateExistingElement = True Then AllCards(FoundIdx) = Card
            End If
            Select Case Card.Keyword
                Case eFITSKeywords.BITPIX : MyProps.BitPix = CInt(Card.Value)
                Case eFITSKeywords.NAXIS : MyProps.NAXIS = CInt(Card.Value)
                Case eFITSKeywords.NAXIS1 : MyProps.Width = CInt(Card.Value)
                Case eFITSKeywords.NAXIS2 : MyProps.Height = CInt(Card.Value)
                Case eFITSKeywords.NAXIS3 : MyProps.NAXIS3 = CInt(Card.Value)
                Case eFITSKeywords.BZERO : MyProps.BZERO = CDbl(Card.Value)
                Case eFITSKeywords.BSCALE : MyProps.BSCALE = CDbl(Card.Value)
            End Select
        Next Card
    End Sub

    '''<summary>Add the card to the already existing list of cards.</summary>
    Public Sub Add(ByVal CardToAdd As cFITSHeaderParser.sHeaderElement)
        'Check and update found element is exists
        Dim FoundIdx As Integer = IdxOfKeyword(CardToAdd.Keyword)
        If FoundIdx = -1 Then
            AllCards.Add(CardToAdd)
        Else
            If UpdateExistingElement = True Then AllCards(FoundIdx) = CardToAdd
        End If
        'Set property entries again
        Select Case CardToAdd.Keyword
            Case eFITSKeywords.BITPIX : MyProps.BitPix = CInt(ElementValue(CardToAdd.Keyword))
            Case eFITSKeywords.NAXIS1 : MyProps.Width = CInt(ElementValue(CardToAdd.Keyword))
            Case eFITSKeywords.NAXIS2 : MyProps.Height = CInt(ElementValue(CardToAdd.Keyword))
            Case eFITSKeywords.NAXIS3 : MyProps.NAXIS3 = CInt(ElementValue(CardToAdd.Keyword))
            Case eFITSKeywords.BZERO : MyProps.BZERO = CDbl(ElementValue(CardToAdd.Keyword))
            Case eFITSKeywords.BSCALE : MyProps.BSCALE = CDbl(ElementValue(CardToAdd.Keyword))
        End Select
    End Sub

    '''<summary>Return the value of the given keyword if present.</summary>
    '''<returns>Index of the found element or -1 if element is not found.</returns>
    Public Function ElementValue(ByVal Keyword As eFITSKeywords) As Object
        For Idx As Integer = 0 To AllCards.Count - 1
            If AllCards(Idx).Keyword = Keyword Then Return AllCards(Idx).Value
        Next Idx
        Return Nothing
    End Function

    '''<summary>Check if an element with the given keyword already exists in the list of elements.</summary>
    '''<returns>Index of the found element or -1 if element is not found.</returns>
    Private Function IdxOfKeyword(ByVal Keyword As eFITSKeywords) As Integer
        For Idx As Integer = 0 To AllCards.Count - 1
            If AllCards(Idx).Keyword = Keyword Then Return Idx
        Next Idx
        Return -1
    End Function

    '''<summary>Get a keyword-value dictionary.</summary>
    '''<remarks>If an entry is found again, the latest present entry in the list will be returned.</remarks>
    Public Function GetCardsAsDictionary() As Dictionary(Of eFITSKeywords, Object)
        Dim RetVal As New Dictionary(Of eFITSKeywords, Object)
        For Each Entry As cFITSHeaderParser.sHeaderElement In AllCards
            Dim Keyword As eFITSKeywords = Entry.Keyword
            If RetVal.ContainsKey(Keyword) = False Then
                RetVal.Add(Keyword, Entry.Value)          'entry is new -> add
            Else
                RetVal(Keyword) = Entry.Value             'entry already exists -> update
            End If
        Next Entry
        Return RetVal
    End Function

    '''<summary>Get a keyword-value dictionary.</summary>
    '''<remarks>If an entry is found again, the latest present entry in the list will be returned.</remarks>
    Public Shared Function GetLCardsAsDictionary(ByRef CardsToProcess As List(Of cFITSHeaderParser.sHeaderElement)) As Dictionary(Of eFITSKeywords, Object)
        Dim RetVal As New Dictionary(Of eFITSKeywords, Object)
        For Each Card As cFITSHeaderParser.sHeaderElement In CardsToProcess
            Dim Keyword As eFITSKeywords = Card.Keyword
            If RetVal.ContainsKey(Keyword) = False Then
                RetVal.Add(Keyword, Card.Value)          'entry is new -> add
            Else
                RetVal(Keyword) = Card.Value             'entry already exists -> update
            End If
        Next Card
        Return RetVal
    End Function

    '''<summary>Get all cards as list.</summary>
    '''<remarks>If an entry is found again, the latest present entry in the list will be returned.</remarks>
    Public Function GetCardsAsList() As Dictionary(Of eFITSKeywords, Object)
        Dim RetVal As New Dictionary(Of eFITSKeywords, Object)
        For Each Entry As cFITSHeaderParser.sHeaderElement In AllCards
            Dim Keyword As eFITSKeywords = Entry.Keyword
            If RetVal.ContainsKey(Keyword) = False Then
                RetVal.Add(Keyword, Entry.Value)
            Else
                RetVal(Keyword) = Entry.Value
            End If
        Next Entry
        Return RetVal
    End Function

    '''<summary>Try to translate the string in the enum.</summary>
    Public Shared Function GetKeywordEnum(ByVal Keyword As String) As eFITSKeywords
        For Each EnumKey As eFITSKeywords In [Enum].GetValues(GetType(eFITSKeywords))
            If EnumKey.ToString.ToUpper = Keyword.ToUpper.Trim Then Return EnumKey
        Next EnumKey
        Return Nothing
    End Function

End Class