Imports System.Net.Mail

Public Class Form1


    Public Function EnviarEmail()


        Dim retFuncao As String = "0"



        Try


            ' Carrega os dados simulando os relatorios resumidos
            Dim dtdados As New DataTable

            dtdados.Columns.Add("codigoRelatorio")
            dtdados.Columns.Add("descricaoRelatorio")
            dtdados.Columns.Add("infoRelatorio")
            dtdados.Columns.Add("valorRelatorio")


            For i As Integer = 1 To 10

                dtdados.Rows.Add(i, "Esse é o relatório " & i, "Todas as lojas", i * 1000)

            Next

            ' FIM: Carrega os dados simulando os relatorios resumidos





            ' Se for por mais de 1 separar por virgula

            Dim destinatario As String = "luizfquintanilha@hotmail.com" 'PODE ALTERAR COM SEU EMAIL OUTLOOK

            Dim Assunto As String = "Aqui está seu resumo de"
            Dim CorpoEmail As String = ""
            Dim DataAtual As String = Today.ToString("dd/MM/yyyy")


            Dim NomeRemetente As String = "Naxos - Resumo Diário" '
            Dim EmailRemetente As String = "<seuGMAIL@gmail.com>" 'PODE ALTERAR COM SEU GMAIL FORMATO: <EMAIL@GMAIL.COM>
            Dim SenhaRemetente As String = "sua senha" 'PODE ALTERAR COM A SENHA DO SEU GMAIL



            ' Primeira linha do email

            CorpoEmail = "Se liga no resumo que montamos para você.. 😉👇🏻" & vbCrLf & vbCrLf


            ' Laco que adiciona cada relatorio no corpo do email

            For J As Integer = 0 To dtdados.Rows.Count - 1

                Dim VALOR As Decimal = dtdados(J)(3).ToString
                Dim STRVALOR As String = VALOR.ToString("R$ #,###.00")
                If VALOR = 0 Then
                    STRVALOR = ("R$ 0,00")
                End If




                CorpoEmail = CorpoEmail & dtdados.Rows(J)(0) & " - " & dtdados.Rows(J)(1) & vbCrLf &
                        "     " & dtdados.Rows(J)(2) & vbCrLf &
                        "     " & STRVALOR & vbCrLf &
                        vbCrLf


            Next



            ' Rodapé

            CorpoEmail = vbCrLf & CorpoEmail & "Isso é tudo... 🤗" & vbCrLf & vbCrLf

            CorpoEmail = vbCrLf & CorpoEmail & "Caso tenha alguma dica ou sugestão 🧐 para esse relatório, responda esse e-mail com sua solicitação e iremos avaliar com carinho 👨‍💻" & vbCrLf











            ' Envia

            Dim mail As New MailMessage
            Dim smtp As New SmtpClient("smtp.gmail.com")

            mail.From = New MailAddress(NomeRemetente & EmailRemetente)
            mail.To.Add(destinatario)
            mail.Subject = Assunto & " " & DataAtual & " 🤩🤑"
            mail.IsBodyHtml = True
            mail.Body = (CorpoEmail)

            smtp.Port = 587
            smtp.UseDefaultCredentials = True
            smtp.Credentials = New System.Net.NetworkCredential(EmailRemetente.Replace("<", "").Replace(">", ""), SenhaRemetente)
            smtp.EnableSsl = True
            smtp.Send(mail)



            System.Threading.Thread.Sleep(30000)
            Application.DoEvents()




            retFuncao = "1"




        Catch ex As Exception

            MsgBox(retFuncao & " " & ex.Message)

        End Try








        Return retFuncao

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If EnviarEmail() = "1" Then
            Close()

        End If

    End Sub
End Class
