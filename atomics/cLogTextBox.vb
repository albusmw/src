Option Explicit On
Option Strict On

Public Class cLogTextBox : Inherits TextBox

    Private TextBuilder As New System.Text.StringBuilder

    Public Sub New()
        MyBase.Multiline = True
        MyBase.ScrollBars = ScrollBars.Both
        MyBase.ReadOnly = True
        MyBase.Font = New Font("Courier New", 8)
    End Sub

    Public Property TimeStampFormat() As String = "yyyy-MM-dd hh:nn:ss"

    Public Sub Log(ByVal Text As String)
        TextBuilder.AppendLine(Text)
        MyBase.Text = TextBuilder.ToString
    End Sub

End Class
