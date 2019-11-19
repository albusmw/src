Option Explicit On
Option Strict On

'''<summary>This class is used to generate Richt Text Font (RTF) output.</summary>
Public Class cRTFGenerator

    '''<summary>Structure containing all information to generate an RTF font object.</summary>
    Public Structure sRTFFontCode

        Public FontSize As Integer
        Public Bold As Boolean
        Public Italic As Boolean
        Public Underline As Boolean
        Public Alignment As eTextAlignment

        Public Sub New(ByVal FontSize As Integer, ByVal Bold As Boolean, ByVal Italic As Boolean, ByVal Underline As Boolean, ByVal Alignment As eTextAlignment)
            Me.FontSize = FontSize
            Me.Bold = Bold
            Me.Italic = Italic
            Me.Underline = Underline
            Me.Alignment = Alignment
        End Sub

        Public Function BuildRTFCodeForFontAndAlignment() As String

            Dim RetVal As String = String.Empty

            'Generate format code
            RetVal &= "\fs" & (FontSize * 2)
            If Bold = True Then RetVal &= "\b"
            If Italic = True Then RetVal &= "\i"
            If Underline = True Then RetVal &= "\ul"
            Select Case Alignment
                Case eTextAlignment.Left : RetVal &= "\ql"
                Case eTextAlignment.Right : RetVal &= "\qr"
                Case eTextAlignment.Middle : RetVal &= "\qc"
            End Select

            Return RetVal

        End Function

    End Structure

    '''<summary>Buffer that contains the RTF text to display.</summary>
    Private RTFText As New Text.StringBuilder
    '''<summary>List of available RTF text / background colors.</summary>
    Private Colors As Dictionary(Of Drawing.Color, Integer)
    '''<summary>Color table in the RTF document.</summary>
    Private ColorRTF As String = "{\colortbl ;"

    Public Enum eTextAlignment
        Left = 0
        Middle = 1
        Right = 2
    End Enum

    '====================================================================================================
    'PUBLIC PROPERTIES
    '====================================================================================================

    '''<summary>Refresh the associated control if a new text arrives.</summary>
    Public Property AutoRefresh() As Boolean
        Get
            Return MyAutoRefresh
        End Get
        Set(ByVal value As Boolean)
            MyAutoRefresh = value
            If value = True Then RefreshRTF()
        End Set
    End Property
    Private MyAutoRefresh As Boolean = False

    '''<summary>Automatically scroll to the last line of the text.</summary>
    Public Property AutoScroll() As Boolean
        Get
            Return MyAutoScroll
        End Get
        Set(ByVal value As Boolean)
            MyAutoScroll = value
        End Set
    End Property
    Private MyAutoScroll As Boolean = True

    '''<summary>Font type to use if no font type is specified.</summary>
    Public Property DefaultFontName() As String
        Get
            Return MyDefaultFontName
        End Get
        Set(ByVal value As String)
            MyDefaultFontName = value
        End Set
    End Property
    Private MyDefaultFontName As String = "Courier New"

    '''<summary>Font size to use if no special font size is specified.</summary>
    Public Property DefaultFontSize() As Integer
        Get
            Return MyDefaultFontSize
        End Get
        Set(ByVal value As Integer)
            MyDefaultFontSize = value
        End Set
    End Property
    Private MyDefaultFontSize As Integer = 16

    '====================================================================================================
    'PUBLIC METHODS
    '====================================================================================================

    '''<summary>Attach this RTF text generator to a RichTextBox control.</summary>
    '''<param name="RTFTextBox">RichTextBox to attach to.</param>
    Public Sub AttachToControl(ByRef RTFTextBox As Windows.Forms.RichTextBox)
        MyRTFBox = RTFTextBox
    End Sub
    Private MyRTFBox As Windows.Forms.RichTextBox

    '''<summary>Init the passed RTF box.</summary>
    Public Sub RTFInit()
        RTFInit(DefaultFontName, DefaultFontSize)
    End Sub

    '''<summary>Init the passed RTF box.</summary>
    Public Sub RTFInit(ByVal NewDefaultFontSize As Integer)
        DefaultFontSize = NewDefaultFontSize
        RTFInit(DefaultFontName, NewDefaultFontSize)
    End Sub

    '''<summary>Init the passed RTF box.</summary>
    '''<param name="NewDefaultFontName">Font name to select.</param>
    '''<param name="NewDefaultFontSize">Font size to select.</param>
    Public Sub RTFInit(ByVal NewDefaultFontName As String, ByVal NewDefaultFontSize As Integer)
        DefaultFontName = NewDefaultFontName
        DefaultFontSize = NewDefaultFontSize
        MyRTFBox.Rtf = "{" + FormatHeader + ColorRTF + "}"
        If AutoRefresh = True Then RefreshRTF()
    End Sub

    Public Sub Clear()
        RTFText = New Text.StringBuilder
        RefreshRTF()
    End Sub

    ''' <summary>Add a line to the RTF box.</summary>
    ''' <param name="Text">Text to be added.</param>
    ''' <param name="ForeColor">Fore color to use.</param>
    ''' <param name="NewLine">Append a new line?</param>
    ''' <param name="Bold">Format bold?</param>
    ''' <param name="Italic">Format italic?</param>
    ''' <param name="Underline">Format underligned?</param>
    ''' <param name="Alig">Alignment of the text.</param>
    Public Sub AddEntry(ByVal Text As String, ByVal ForeColor As Drawing.Color, Optional ByVal NewLine As Boolean = True, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal Underline As Boolean = False, Optional ByVal Alig As eTextAlignment = eTextAlignment.Left)
        AddEntry(Text, ForeColor, False, Nothing, NewLine, Bold, Italic, Underline, Alig)
    End Sub


    '''<summary>Add a line to the RTF box.</summary>
    '''<param name="Text">Text to be added.</param>
    '''<param name="ForeColor">Fore color to use.</param>
    '''<param name="BackColor">Back color to apply.</param>
    '''<param name="NewLine">Append a new line?</param>
    '''<param name="Bold">Format bold?</param>
    '''<param name="Italic">Format italic?</param>
    '''<param name="Underline">Format underligned?</param>
    '''<param name="Alig">Alignment of the text.</param>
    Public Sub AddEntry(ByVal Text As String, ByVal ForeColor As Drawing.Color, ByVal ApplyBackColor As Boolean, ByVal BackColor As Drawing.Color, Optional ByVal NewLine As Boolean = True, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal Underline As Boolean = False, Optional ByVal Alig As eTextAlignment = eTextAlignment.Left)

        Dim FormatCode As String

        'Basic format code
        FormatCode = New sRTFFontCode(DefaultFontSize, Bold, Italic, Underline, Alig).BuildRTFCodeForFontAndAlignment

        'Create new color entries if required and add color coding
        If IsNothing(ForeColor) = False Then
            FormatCode &= "\cf" & (AddColor(ForeColor) + 1).ToString.Trim
        End If
        If ApplyBackColor Then
            FormatCode &= "\highlight" & (AddColor(ForeColor) + 1).ToString.Trim
        End If

        'Set font
        FormatCode &= CreateFontType()

        '==============================================================

        AddToRTF(Text, FormatCode, NewLine)

    End Sub

    Public Sub ForceRefresh()
        RefreshRTF()
    End Sub

    '====================================================================================================
    'HELPER FUNCTIONS
    '====================================================================================================

    Private ReadOnly Property FontTable() As String
        Get
            Dim RetVal As String = "{\fonttbl"
            RetVal &= "{\f0\fmodern\fprq1\fcharset0\fs" & DefaultFontSize.ToString.Trim & " " & DefaultFontName & ";}"
            RetVal &= "{\f1\fswiss\fprq2\fcharset0 Microsoft Sans Serif;}"
            RetVal &= "{\f2\fswiss\fcharset0 Arial;}"
            RetVal &= "}"
            Return RetVal
        End Get
    End Property

    ''' <summary>Basic formating rule.</summary>
    ''' <returns>String containing the basic formating rule.</returns>
    Private ReadOnly Property FormatHeader() As String
        Get
            ' \plain\fs" & CStr(FontSize) &
            Return "\rtf1" & FontTable & ColorRTF
        End Get
    End Property

    Private Function CreateFontType() As String
        Return "\f0"
    End Function

    '''<summary>Add the text, formated with the specific format code, to the associated RTF box.</summary>
    '''<param name="Text">Text to add - if the text is nothing, no further actions are performed.</param>
    '''<param name="FormatCode">Format generated by the AddEntry code.</param>
    '''<param name="NewLine">Move to a new line after the command.</param>
    Private Sub AddToRTF(ByVal Text As String, ByVal FormatCode As String, ByVal NewLine As Boolean)
        If IsNothing(Text) = False Then
            With MyRTFBox
                Text = Text.Replace("\", "\\")
                If NewLine = True Then Text &= "\par"
                RTFText.Append("{" + FormatCode + " " + Text + "}")
                If AutoRefresh Then RefreshRTF()
            End With
        End If
    End Sub

    '''<summary>Add a color to the RTF color table.</summary>
    '''<param name="NewColor">Color to add to the list.</param>
    Private Function AddColor(ByVal NewColor As Drawing.Color) As Integer
        'If the color is missing
        If Colors.ContainsKey(NewColor) = False Then
            'Add the color
            Colors.Add(NewColor, Colors.Count)
            'Generate a new list for thr ColorRTF field
            Dim AllColors As Dictionary(Of Drawing.Color, Integer).KeyCollection = Colors.Keys
            ColorRTF = String.Empty
            For Each SingleColor As Drawing.Color In AllColors
                ColorRTF &= "\red" & SingleColor.R.ToString.Trim & "\green" & SingleColor.G.ToString.Trim & "\blue" & SingleColor.B.ToString.Trim & ";"
            Next SingleColor
            ColorRTF = "{\colortbl ;" & ColorRTF & "}"
        End If
        If Colors.ContainsKey(NewColor) Then Return Colors.Item(NewColor) Else Return 0
    End Function

    Public Function GetRTFText() As String
        Return "{" + FormatHeader + RTFText.ToString + "}"
    End Function

    '''<summary>Refresh the associated RTF box.</summary>
    '''<todo>Check what to do with font size.</todo>
    Private Sub RefreshRTF()
        If IsNothing(MyRTFBox) = False Then
            With MyRTFBox
                .Rtf = "{" + FormatHeader + RTFText.ToString + "}"
                .SelectionStart = .Text.Length
                If AutoScroll = True Then .ScrollToCaret()
                '.Refresh()                 'not required
            End With
        End If
    End Sub

    Public Sub New()
        Colors = New Dictionary(Of Drawing.Color, Integer)
    End Sub

End Class