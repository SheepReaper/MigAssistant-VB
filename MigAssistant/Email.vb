Imports System.Net.Mail
Imports System.Text

Public Class Email
    Private _strMailServer As String
    Private _strMailRecipients As String
    Private _strMailFrom As String
    Private _strMailSubject As String
    Private _strMailMessage As String
    Private _strMailAttachments As String

    Public Property Server As String
        Get
            Return _strMailServer
        End Get
        Set
            _strMailServer = Value
        End Set
    End Property

    Public Property Recipients As String
        Get
            Return _strMailRecipients
        End Get
        Set
            _strMailRecipients = Value
        End Set
    End Property

    Public Property From As String
        Get
            Return _strMailFrom
        End Get
        Set
            _strMailFrom = Value
        End Set
    End Property

    Public Property Subject As String
        Get
            Return _strMailSubject
        End Get
        Set
            _strMailSubject = Value
        End Set
    End Property

    Public Property Message As String
        Get
            Return _strMailMessage
        End Get
        Set
            _strMailMessage = Value
        End Set
    End Property

    Public Property Attachments As String
        Get
            Return _strMailAttachments
        End Get
        Set
            _strMailAttachments = Value
        End Set
    End Property

    Public Sub Send()
        'This procedure takes string array parameters for multiple recipients and files
        Try
            'For each to address create a mail message
            Dim mailMessage As New MailMessage With {
                .BodyEncoding = Encoding.Default,
                .Subject = _strMailSubject.Trim(),
                .Body = _strMailMessage.Trim() & vbCrLf,
                .From = New MailAddress(_strMailFrom.Trim()),
                .Priority = MailPriority.Normal,
                .IsBodyHtml = True
            }

            For Each recipient As String In Split(_strMailRecipients, ",")
                mailMessage.To.Add(New MailAddress(recipient.Trim()))
            Next

            'attach each file attachment
            For Each attachment As String In Split(_strMailAttachments, ",")
                If Not attachment = "" Or Nothing Then
                    Dim msgAttach As New Attachment(attachment)
                    mailMessage.Attachments.Add(msgAttach)
                End If
            Next

            'Smtpclient to send the mail message
            Dim smtpMail As New SmtpClient With {
                .Host = _strMailServer
            }
            smtpMail.Send(mailMessage)
            'Message Successful
        Catch exSmtpFailedRecipients As SmtpFailedRecipientsException
            Throw New Exception(exSmtpFailedRecipients.Message)
        Catch exSmtpFailedRecipient As SmtpFailedRecipientException
            Throw New Exception(exSmtpFailedRecipient.Message)
        Catch exSmtp As SmtpException
            Throw New Exception(exSmtp.Message)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

End Class