Imports System.Net.Mail
Module PasivosIntereses

    Sub GeneraInteresesDiarios(FechaFin As Date)
        Dim TaTasas As New ProduccionDSTableAdapters.HistaTableAdapter
        Dim FecIni, FecAux As Date
        Dim SaldoIni, Tasa, Dias, Interes, Retencion, Capital As Decimal
        Dim FondeoDS As New WEB_FinagilDS
        Dim taFond As New WEB_FinagilDSTableAdapters.DatosFondeosTableAdapter
        Dim taEdoCta As New WEB_FinagilDSTableAdapters.FOND_EstadoCuentaTableAdapter
        Dim Mov As WEB_FinagilDS.FOND_EstadoCuentaRow
        Dim taPag As New WEB_FinagilDSTableAdapters.FOND_FechasPagoCapitalTableAdapter
        Dim Pago As WEB_FinagilDS.FOND_FechasPagoCapitalRow
        taFond.Fill(FondeoDS.DatosFondeos)
        For Each F As WEB_FinagilDS.DatosFondeosRow In FondeoDS.DatosFondeos.Rows
            SaldoIni = 0
            If F.id_Fondeo <> 20 Then
                'Continue For
            End If
            If F.Tipo_Fondeo = "INDIVIDUAL" Then
                taEdoCta.QuitarInteresPagoAut(F.id_Fondeo, FechaFin.Month, FechaFin.Year)
                SaldoIni = taEdoCta.SumaCapital(F.id_Fondeo)
                FecAux = taEdoCta.UltimaFechaFin(F.id_Fondeo)
                FecIni = FecAux
                While FecAux <= FechaFin
                    If F.TipoTasa = "Tasa Fija" Then
                        Tasa = F.TasaDiferencial
                        If Tasa = 0 Then ErrorEnTasa("atorres@finagil.com.mx", F.TipoTasa, FecAux)
                    End If
                    taPag.Fill(FondeoDS.FOND_FechasPagoCapital, F.id_Fondeo, FecAux)
                    If FondeoDS.FOND_FechasPagoCapital.Rows.Count > 0 And FecAux.Month = FechaFin.Month Then
                        Pago = FondeoDS.FOND_FechasPagoCapital.Rows(0)
                        Dias = DateDiff(DateInterval.Day, FecIni, FecAux)
                        Interes = (Tasa / 36000) * Dias * SaldoIni
                        Retencion = Math.Round(SaldoIni * Math.Round(F.TasaRetencion / 36000, 6), 2)
                        taEdoCta.Insert(F.id_Fondeo, "INTERESES", 0, Interes, Retencion, F.TasaRetencion, FecIni, FecAux, SaldoIni, SaldoIni)
                        Interes = taEdoCta.SumaInteres(F.id_Fondeo) * -1
                        taEdoCta.Insert(F.id_Fondeo, "PAGO AUTOMATICO", Pago.Capital * -1, Interes, Retencion, F.TasaRetencion, FecIni, FecAux, SaldoIni, SaldoIni - Pago.Capital)
                        SaldoIni -= Pago.Capital
                        FecIni = FecAux
                    End If
                    FecAux = FecAux.AddDays(1)
                End While
                'corte de interes a la fecha***********************************
                Dias = DateDiff(DateInterval.Day, FecIni, FechaFin)
                Interes = (Tasa / 36000) * Dias * SaldoIni
                Retencion = Math.Round(SaldoIni * Math.Round(F.TasaRetencion / 36000.6), 2)
                If Dias > 1 Then
                    taEdoCta.Insert(F.id_Fondeo, "INTERESES", 0, Interes, Retencion, F.TasaRetencion, FecIni, FechaFin, SaldoIni, SaldoIni)
                End If
                '**corte de interes a la fecha***********************************

            Else
                FecIni = FechaFin.AddDays((FechaFin.Day - 1) * -1)
                taEdoCta.QuitaInteresesMes(F.id_Fondeo, FechaFin.Month, FechaFin.Year)
                While FecIni <= FechaFin
                    If F.TipoTasa = "Tasa Fija" Then
                        Tasa = F.TasaDiferencial
                    ElseIf F.TipoTasa = "Tasa TIIE 28" Then
                        Tasa = F.TasaDiferencial + TaTasas.SacaTASA(4, FecIni.ToString("yyyyMMdd"))
                        If Tasa = 0 Then ErrorEnTasa("atorres@finagil.com.mx", F.TipoTasa, FecIni)
                    ElseIf F.TipoTasa = "Tasa Libor" Then
                        Tasa = F.TasaDiferencial + TaTasas.SacaTASA(12, FecIni.ToString("yyyyMMdd"))
                        For x As Integer = 1 To 4
                            Tasa = F.TasaDiferencial + TaTasas.SacaTASA(12, FecIni.AddDays(x * -1).ToString("yyyyMMdd"))
                            If Tasa > 0 Then Exit For
                        Next
                        If Tasa = 0 Then
                            ErrorEnTasa("atorres@finagil.com.mx", F.TipoTasa, FecIni)
                        End If
                    End If

                    If SaldoIni = 0 Then
                        SaldoIni = taEdoCta.SumaCapitalHasta(F.id_Fondeo, FecIni.AddDays(-1))
                    End If
                    Capital = taEdoCta.SacaCAPITAL(FecIni, F.id_Fondeo)

                    If SaldoIni > 0 And Capital = 0 Then
                        Mov = FondeoDS.FOND_EstadoCuenta.NewFOND_EstadoCuentaRow
                        Mov.Concepto = "INTERESES"
                        Mov.id_Fondeo = F.id_Fondeo
                        Mov.SaldoInicial = SaldoIni
                        Mov.Interes = SaldoIni * (Tasa / 36000)
                        Mov.Retencion = Math.Round(SaldoIni * Math.Round(F.TasaRetencion / 36000, 6), 2)
                        Mov.SaldoFinal = Mov.SaldoInicial
                        Mov.FechaInicio = FecIni
                        Mov.FechaFin = FecIni
                        Mov.Importe = 0
                        Mov.TasaRetencion = F.TasaRetencion
                        Mov.EndEdit()
                        FondeoDS.FOND_EstadoCuenta.AddFOND_EstadoCuentaRow(Mov)
                        taEdoCta.Update(FondeoDS.FOND_EstadoCuenta)
                    ElseIf Capital = 0 And SaldoIni = 0 Then
                    Else
                        taEdoCta.FillByFecha(FondeoDS.FOND_EstadoCuenta, F.id_Fondeo, FecIni)
                        Mov = FondeoDS.FOND_EstadoCuenta.Rows(0)
                        Mov.BeginEdit()
                        SaldoIni += Capital
                        Mov.SaldoInicial = SaldoIni
                        Mov.Interes = Mov.SaldoInicial * (Tasa / 36000)
                        Mov.Retencion = Math.Round(Mov.SaldoInicial * Math.Round(F.TasaRetencion / 36000, 6), 2)
                        Mov.SaldoFinal = SaldoIni + Mov.Importe
                        Mov.FechaInicio = FecIni
                        Mov.TasaRetencion = F.TasaRetencion
                        Mov.FechaFin = FecIni
                        Mov.EndEdit()
                        taEdoCta.Update(Mov)
                    End If
                    FecIni = FecIni.AddDays(1)
                End While
            End If


        Next

    End Sub

    Sub ErrorEnTasa(Para As String, Tasa As String, Fecha As Date)
        Try
            Dim Servidor As New SmtpClient
            Dim Mensaje As New MailMessage
            Servidor.Host = "smtp01.cmoderna.com"
            Servidor.Port = "26"
            Mensaje.To.Add(Para)
            Mensaje.To.Add("ecacerest@finagil.com.mx")
            Mensaje.From = New MailAddress("Tasas@Finagil.com.mx", "FINAGIL envíos automáticos")
            Mensaje.Subject = "Tasa no Encontrada"
            Mensaje.Body = "Tasa no encontrada: " & Tasa & "<BR> Fecha : " & Fecha.ToShortDateString
            Servidor.Send(Mensaje)
        Catch ex As Exception

        End Try
    End Sub

End Module
