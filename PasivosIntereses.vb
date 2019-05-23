Imports System.Net.Mail
Module PasivosIntereses
    Dim CorreoTESORERIA As String = "atorres@finagil.com.mx"
    Sub GeneraInteresesDiarios(FechaFin As Date, idAux As Integer)
        Console.WriteLine("ID=" & idAux & " - " & FechaFin.ToShortDateString)
        Dim TaTasas As New ProduccionDSTableAdapters.HistaTableAdapter
        Dim FecIni, FecAux As Date
        Dim SaldoIni, Tasa, Dias, Interes, Retencion, Capital As Decimal
        Dim FondeoDS As New WEB_FinagilDS
        Dim taFond As New WEB_FinagilDSTableAdapters.DatosFondeosTableAdapter
        Dim taEdoCta As New WEB_FinagilDSTableAdapters.FOND_EstadoCuentaTableAdapter
        Dim Mov As WEB_FinagilDS.FOND_EstadoCuentaRow
        Dim taPag As New WEB_FinagilDSTableAdapters.FOND_FechasPagoCapitalTableAdapter
        Dim Pago As WEB_FinagilDS.FOND_FechasPagoCapitalRow
        Dim TasaRetencion As Decimal = 0

        If idAux <> 0 Then
            taFond.FillByIdFondeo(FondeoDS.DatosFondeos, idAux)
        Else
            taFond.Fill(FondeoDS.DatosFondeos)
        End If

        For Each F As WEB_FinagilDS.DatosFondeosRow In FondeoDS.DatosFondeos.Rows
            SaldoIni = 0

            If F.Tipo_Fondeo = "INDIVIDUAL" Then
                taEdoCta.QuitarInteresPagoAut(F.id_Fondeo, FechaFin.Month, FechaFin.Year)
                SaldoIni = taEdoCta.SumaCapital(F.id_Fondeo)
                FecAux = taEdoCta.UltimaFechaFin(F.id_Fondeo)
                FecIni = FecAux
                While FecAux <= FechaFin
                    TasaRetencion = taFond.SacaTasaRetension(FecAux)
                    If F.TasaRetencion = 0 Then
                        TasaRetencion = 0
                    End If
                    If F.TipoTasa = "Tasa Fija" Then
                        Tasa = F.TasaDiferencial
                        If Tasa = 0 Then ErrorEnTasa(CorreoTESORERIA, F.TipoTasa, FecAux)
                    End If
                    taPag.Fill(FondeoDS.FOND_FechasPagoCapital, F.id_Fondeo, FecAux)
                    If FondeoDS.FOND_FechasPagoCapital.Rows.Count > 0 And FecAux.Month = FechaFin.Month Then
                        Pago = FondeoDS.FOND_FechasPagoCapital.Rows(0)
                        If FecAux.DayOfWeek = DayOfWeek.Sunday Then
                            FecAux = FecAux.AddDays(1)
                            taPag.UpdateFecha(FecAux, Pago.id_pago)
                        End If
                        If FecAux.DayOfWeek = DayOfWeek.Saturday Then
                            FecAux = FecAux.AddDays(2)
                            taPag.UpdateFecha(FecAux, Pago.id_pago)
                        End If
                        Dias = DateDiff(DateInterval.Day, FecIni, FecAux)
                        Interes = (Tasa / 36000) * Dias * SaldoIni
                        Retencion = Math.Round(SaldoIni * Math.Round(TasaRetencion / 360, 6), 2)
                        taEdoCta.Insert(F.id_Fondeo, "INTERESES", 0, Interes, Retencion, TasaRetencion, FecIni, FecAux, SaldoIni, SaldoIni, 0, "")
                        Interes = taEdoCta.SumaInteres(F.id_Fondeo) * -1
                        taEdoCta.Insert(F.id_Fondeo, "PAGO AUTOMATICO", Pago.Capital * -1, Interes, Retencion, TasaRetencion, FecAux, FecAux, SaldoIni, SaldoIni - Pago.Capital, 0, F.BancoDefault)
                        SaldoIni -= Pago.Capital
                        FecIni = FecAux
                    ElseIf FecAux.Month <> FechaFin.Month Then ' corte de interes
                        taEdoCta.QuitaCorteInteres(F.id_Fondeo, FecIni, FecAux)
                        Dias = DateDiff(DateInterval.Day, FecIni, FecAux.AddDays(1))
                        Interes = (Tasa / 36000) * Dias * SaldoIni
                        Retencion = Math.Round(SaldoIni * Math.Round(TasaRetencion / 360, 6), 2)
                        taEdoCta.Insert(F.id_Fondeo, "INTERESES", 0, Interes, Retencion, TasaRetencion, FecIni, FecAux, SaldoIni, SaldoIni, 0, "")
                        FecIni = FecAux.AddDays(1)
                    End If
                    FecAux = FecAux.AddDays(1)
                End While
                'corte de interes a la fecha***********************************
                Dias = DateDiff(DateInterval.Day, FecIni, FechaFin)
                Interes = (Tasa / 36000) * Dias * SaldoIni
                Retencion = Math.Round(SaldoIni * Math.Round(TasaRetencion / 360), 2)
                If Dias > 1 Then
                    taEdoCta.Insert(F.id_Fondeo, "INTERESES", 0, Interes, Retencion, TasaRetencion, FecIni, FechaFin, SaldoIni, SaldoIni, 0, "")
                End If
                '**corte de interes a la fecha***********************************

            Else
                FecIni = FechaFin.AddDays((FechaFin.Day - 1) * -1)
                taEdoCta.QuitaInteresesMes(F.id_Fondeo, FechaFin.Month, FechaFin.Year)
                While FecIni <= FechaFin
                    TasaRetencion = taFond.SacaTasaRetension(FecIni)
                    If F.TasaRetencion = 0 Then
                        TasaRetencion = 0
                    End If
                    If FecIni > F.FechaVencimiento Then
                        Exit While
                    End If
                    If F.TipoTasa = "Tasa Fija" Then
                        Tasa = F.TasaDiferencial
                    ElseIf F.TipoTasa = "Tasa TIIE 28" Then
                        Tasa = F.TasaDiferencial + TaTasas.SacaTASA(4, FecIni.ToString("yyyyMMdd"))
                        If Tasa = 0 Then ErrorEnTasa(CorreoTESORERIA, F.TipoTasa, FecIni)
                    ElseIf F.TipoTasa = "Tasa Libor" Then
                        Tasa = F.TasaDiferencial + TaTasas.SacaTASA(12, FecIni.ToString("yyyyMMdd"))
                        For x As Integer = 1 To 4
                            Tasa = F.TasaDiferencial + TaTasas.SacaTASA(12, FecIni.AddDays(x * -1).ToString("yyyyMMdd"))
                            If Tasa > 0 Then Exit For
                        Next
                        If Tasa = 0 Then
                            ErrorEnTasa(CorreoTESORERIA, F.TipoTasa, FecIni)
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
                        Mov.Retencion = Math.Round(SaldoIni * Math.Round(TasaRetencion / 360, 6), 2)
                        Mov.SaldoFinal = Mov.SaldoInicial
                        Mov.FechaInicio = FecIni
                        Mov.FechaFin = FecIni
                        Mov.Importe = 0
                        Mov.TasaRetencion = TasaRetencion
                        Mov.EndEdit()
                        FondeoDS.FOND_EstadoCuenta.AddFOND_EstadoCuentaRow(Mov)
                        taEdoCta.Update(FondeoDS.FOND_EstadoCuenta)
                    ElseIf Capital = 0 And SaldoIni = 0 Then
                    Else
                        taEdoCta.FillByFecha(FondeoDS.FOND_EstadoCuenta, F.id_Fondeo, FecIni)
                        Mov = FondeoDS.FOND_EstadoCuenta.Rows(0)
                        Mov.BeginEdit()
                        Mov.SaldoInicial = SaldoIni
                        Mov.SaldoFinal = Mov.SaldoInicial + Mov.Importe
                        If TasaRetencion > 0 Then ' personas morales
                            Mov.Interes = Mov.SaldoFinal * (Tasa / 36000)
                            Mov.Retencion = Math.Round(Mov.SaldoFinal * Math.Round(TasaRetencion / 360, 6), 2)
                        Else ' Bancarios
                            If Mov.Interes < 0 Then ' es pago
                            Else
                                Mov.Interes = Mov.SaldoFinal * (Tasa / 36000)
                            End If
                            Mov.Retencion = 0
                        End If
                        Mov.FechaInicio = FecIni
                        Mov.TasaRetencion = TasaRetencion
                        Mov.FechaFin = FecIni
                        Mov.EndEdit()
                        SaldoIni += Capital
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
            Servidor.Host = My.Settings.SMTP
            Servidor.Port = My.Settings.SMTP_port
            Dim Cred() As String = My.Settings.SMTP_creden.Split(",")
            Servidor.Credentials = New System.Net.NetworkCredential(Cred(0), Cred(1), Cred(2))
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
