Option Explicit On
Option Strict On

'''<summary>Display a simple RTF textbox window.</summary>
Public Class cRTFTextBox

    '''<summary>The form that shall be displayed.</summary>
    Public Hoster As System.Windows.Forms.Form = Nothing
    '''<summary>The RTF text box control inside the form.</summary>
    Private rtfTB As RichTextBox = Nothing

    '''<summary>Prepare.</summary>
    Public Sub Init()
        If IsNothing(Hoster) = True Then Hoster = New System.Windows.Forms.Form
        If IsNothing(rtfTB) = True Then
            rtfTB = New RichTextBox
            Hoster.Controls.Add(rtfTB)
            rtfTB.Dock = Windows.Forms.DockStyle.Fill
            rtfTB.ScrollBars = RichTextBoxScrollBars.Both
        End If
    End Sub

    '''<summary>Display text.</summary>
    Public Sub ShowText(ByVal RTFContent As String)
        Init()
        rtfTB.Rtf = RTFContent
        Hoster.Show()
    End Sub

End Class