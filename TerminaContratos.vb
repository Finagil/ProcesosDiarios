Module TerminaContratos

    Public Sub Termina_Contratos()
        Dim TA As New ProduccionDSTableAdapters.AnexosTerminadosTableAdapter
        Dim T As New ProduccionDS.AnexosTerminadosDataTable
        Dim ta2 As New WEB_FinagilDSTableAdapters.CorreosTableAdapter
        Dim t2 As New WEB_FinagilDS.CorreosDataTable

        TA.Fill(T, Today.AddDays(1).ToString("yyyyMMdd"))
        ta2.Fill(t2, "TERMINACONTRATO")
        For Each R As ProduccionDS.AnexosTerminadosRow In T.Rows
            TA.DesbloqueaAnexo(R.Anexo)
            TA.TerminaContrato(R.Anexo)
            TA.BloqueaAnexo(R.Anexo)
            If R.SaldoFac < 10 Then
                For Each rr As WEB_FinagilDS.CorreosRow In t2.Rows
                    Utilerias.EnviacORREO(rr.Correo, R.AnexoCon, "Terminación de Contrato: " & R.AnexoCon, "Notificaciones@finagil.com.mx")
                Next
            End If
        Next
        'Cancelados por Adelanto a capital
        TA.FillSinSaldoInsoluto(T, Today.AddDays(1).ToString("yyyyMMdd"))
        ta2.Fill(t2, "TERMINACONTRATO")
        For Each R As ProduccionDS.AnexosTerminadosRow In T.Rows
            TA.DesbloqueaAnexo(R.Anexo)
            TA.CancelaContrato(R.Anexo)
            TA.BloqueaAnexo(R.Anexo)
            If R.SaldoFac < 10 Then
                For Each rr As WEB_FinagilDS.CorreosRow In t2.Rows
                    Utilerias.EnviacORREO(rr.Correo, R.AnexoCon, "Cancelación de Contrato: " & R.AnexoCon, "Notificaciones@finagil.com.mx")
                Next
            End If
        Next
        TA.QuitaOpciones()
    End Sub

    Public Sub Terminados_Con_Saldo(fec As Date)
        Dim TA As New ProduccionDSTableAdapters.Vw_TerminadosConSaldoTableAdapter
        Dim T As New ProduccionDS.Vw_TerminadosConSaldoDataTable

        Dim ta2 As New WEB_FinagilDSTableAdapters.CorreosTableAdapter
        Dim t2 As New WEB_FinagilDS.CorreosDataTable
        ta2.Fill(t2, "TERMINACONTRATO")

        TA.Fill(T)
        For Each R As ProduccionDS.Vw_TerminadosConSaldoRow In T.Rows
            TA.TerminaAnexoConSaldo(R.Anexo)
            Utilerias.EnviacORREO("ecacerest@finagil.com.mx", R.Anexo, "Terminación de Contrato con Saldo: " & R.Anexo, "Notificaciones@finagil.com.mx")
        Next

        TA.TermiandosConSaldo_liquidados(T)
        For Each R As ProduccionDS.Vw_TerminadosConSaldoRow In T.Rows
            TA.TerminaAnexo(R.Anexo)
            For Each rr As WEB_FinagilDS.CorreosRow In t2.Rows
                Utilerias.EnviacORREO(rr.Correo, R.Anexo, "Terminación de Contrato (Saldo Pagado): " & R.Anexo, "Notificaciones@finagil.com.mx")
            Next
        Next
        'Termina AV
        TA.TerminaContratosAV_W()
        TA.TerminaContratosConSaldoAV(fec.ToString("yyyyMMdd"))
    End Sub

End Module
