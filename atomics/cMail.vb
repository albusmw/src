Option Explicit On
Option Strict On
Imports System.Reflection

Public Class cMail

    '''<summary>SMTP server to use.</summary>
    Public Property SMTP_server As String = "mail.gmx.net"
    '''<summary>SMTP server port.</summary>
    Public Property SMTP_Port As Integer = 587
    '''<summary>SMTP user name.</summary>
    Public Property SMTP_UserName As String = "albusmw@gmx.de"
    '''<summary>SMTP password.</summary>
    Public Property SMTP_Password As String = ""

    Public Sub Send(ByVal Subject As String, ByVal MessageText As List(Of String), ByVal Attachment As String)
        Send(Subject, MessageText, New List(Of String)({Attachment}))
    End Sub

    Public Sub Send(ByVal Subject As String, ByVal MessageText As List(Of String), ByVal Attachments As List(Of String))

        Dim Mail As New Net.Mail.MailMessage

        With Mail
            .From = New Net.Mail.MailAddress(SMTP_UserName, "MyAstroAlert")
            .To.Add(New Net.Mail.MailAddress(SMTP_UserName))
            .Subject = Subject
            .Body = Join(MessageText.ToArray, System.Environment.NewLine)
            .IsBodyHtml = False
            For Each File As String In Attachments
                .Attachments.Add(New Net.Mail.Attachment(File))
            Next File
        End With

        Using Server As New Net.Mail.SmtpClient(SMTP_server)
            With Server
                .DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                .Port = SMTP_Port
                .Credentials = New System.Net.NetworkCredential(SMTP_UserName, SMTP_Password)
                .EnableSsl = True
                .Send(Mail)
            End With
        End Using

        MessageBox.Show("Mail Send")

    End Sub

End Class