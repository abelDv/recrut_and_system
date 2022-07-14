Imports System.Net.Mail
Imports System.Net

Public Class Email

    Function enviaDB(nome As String, cargo As String, edital As String, email As String, ampla As String, negro As String, pcd As String, andamento As Integer) As String
        Try
            Dim mail As New MailMessage
            Dim smtpServer As New SmtpClient
            Dim erros As String = ""

            smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpServer.UseDefaultCredentials = True
            smtpServer.Credentials = New Net.NetworkCredential("user", "pass")
            smtpServer.Port = 587
            smtpServer.EnableSsl = True
            smtpServer.Host = "smtp.gmail.com"

            mail = New MailMessage
            mail.From = New MailAddress("email")
            mail.[To].Add("email_from") 
            mail.Priority = MailPriority.Normal
            mail.Subject = "Convocação"
            mail.IsBodyHtml = True
            mail.Body = stringEmail(nome, cargo, ampla, negro, pcd, edital)
            smtpServer.Send(mail)

            enviaDB =  "sucesso"

        Catch ex As Exception
            'MsgBox("Erro ao tentar enviar o e-mail - enviaDB: " & ex.Message)
            'enviaDB = enviados
            'lb.Text = lb.Text & vbCr & "Erro ao tentar enviar o e-mail: " & ex.Message
            enviaDB = "Erro ao tentar enviar e-mail para: " & nome & vbCrLf
        End Try

    End Function

    Function stringEmail(nome As String, cargo As String, ampla As String, negros As String, pcd As String, edital As String) As String
        Dim mensagem As String

        mensagem = "mensagem_html_using_parameters"

        Return mensagem
    End Function

End Class
