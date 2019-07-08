Imports System.Net.Mail
Module AviosSaldos
    Dim Servidor As New SmtpClient(My.Settings.SMTP, My.Settings.SMTP_port)
    Dim Credenciales As String() = My.Settings.SMTP_creden.Split(",")
    Dim Mensaje As New MailMessage
    Dim Adjunto As Attachment

    Public Sub SaldosAvios()
        Dim ta As New ProduccionDSTableAdapters.SaldosAviosTableAdapter
        Dim t As New ProduccionDS.SaldosAviosDataTable
        Dim R As ProduccionDS.SaldosAviosRow
        Dim tx As New ProduccionDSTableAdapters.AVI_SaldosTMPTableAdapter
        Console.WriteLine("Poner en Ceros")
        ta.SaldoAceros()
        ta.FillConSaldo(t)
        For Each R In t.Rows
            Console.WriteLine(R.AnexoCon)
            tx.DeleteAnexo(R.Anexo, R.Ciclo)
            tx.Insert(R.Anexo, R.Ciclo, R.Cliente, R.CicloPagare, 0, 0, 0, R.Saldo)
        Next
        ta.FillMinistrado(t)
        'ta.FillMinistradorHasta(t, "20160731")
        For Each R In t.Rows
            Console.WriteLine(R.AnexoCon)
            tx.UpdateMinistrado(R.Imp, R.Fega, R.Garantia, R.Anexo, R.Ciclo, R.Anexo, R.Ciclo)
            tx.UpdateMontoFinanciado(R.Imp + R.Fega + R.Garantia, R.Anexo, R.Ciclo, R.Anexo, R.Ciclo)
        Next
        Aplica_Seguro_Vida()
    End Sub

    Sub Aplica_Seguro_Vida()
        Servidor.Credentials = New System.Net.NetworkCredential(Credenciales(0), Credenciales(1), Credenciales(2))
        Dim ta As New ProduccionDSTableAdapters.VwSegVidaTableAdapter
        Dim t As New ProduccionDS.VwSegVidaDataTable
        Dim R As ProduccionDS.VwSegVidaRow
        Dim taMail As New ProduccionDSTableAdapters.GEN_CorreosFasesTableAdapter
        Dim tmail As New ProduccionDS.GEN_CorreosFasesDataTable
        Dim rr As ProduccionDS.GEN_CorreosFasesRow
        Mensaje.IsBodyHtml = True
        Mensaje.From = New MailAddress("SEGUROSVIDA@Finagil.com.mx", "SEGUROS VIDA envíos automáticos")
        ta.Fill(t)
        taMail.Fill(tmail, "SEGUROSVIDA")
        For Each R In t.Rows
            If R.Tipo = "M" Then
                ta.UpdateSegVida("N", 0, R.Anexo, R.Ciclo)
            Else
                Dim FechaCon As Date = CTOD(R.Fechacon)
                Dim cad As String = R.RFC.Substring(4, 6)
                If CInt(cad.Substring(1, 2)) <= Date.Now.Year - 2000 Then
                    cad = "20" & cad
                Else
                    cad = "19" & cad
                End If
                Dim FechaNac As Date = CTOD(cad)
                Dim Edad As Integer = DateDiff(DateInterval.Year, FechaNac, FechaCon)
                If Edad >= 75 Then
                    ta.UpdateSegVida("N", 0, R.Anexo, R.Ciclo)
                    Mensaje.Subject = "Contrato sin seguro de Vida " & R.AnexoCon
                    Mensaje.Body = "Contrato Sin seguro de Vida por la edad de Cliente: <br>"
                Else
                    ta.UpdateSegVida("S", R.SeguroVida, R.Anexo, R.Ciclo)
                    Mensaje.Subject = "Contrato con seguro de Vida " & R.Anexo
                    Mensaje.Body = "Contrato con seguro de Vida por la edad de Cliente: <br>"
                End If
                For Each rr In tmail.Rows
                    Mensaje.To.Add(Trim(rr.Correo))
                Next
                Mensaje.Body += "Cliente: " & R.Descr & "<br>"
                Mensaje.Body += "Fecha de Nacimiento: " & FechaNac.ToShortDateString & "<br>"
                Mensaje.Body += "Edad: " & Edad & "<br>"
                Servidor.Send(Mensaje)
            End If
        Next
    End Sub
End Module
