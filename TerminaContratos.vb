Module TerminaContratos
    Dim CorreosFases As New ProduccionDSTableAdapters.GEN_CorreosFasesTableAdapter
    Dim TCorreosFases As New ProduccionDS.GEN_CorreosFasesDataTable
    Public Sub Termina_Contratos()
        Dim TA As New ProduccionDSTableAdapters.AnexosTerminadosTableAdapter
        Dim T As New ProduccionDS.AnexosTerminadosDataTable

        TA.Fill(T, Today.AddDays(1).ToString("yyyyMMdd"))
        CorreosFases.Fill(TCorreosFases, "TERMINACONTRATO")
        For Each R As ProduccionDS.AnexosTerminadosRow In T.Rows
            TA.DesbloqueaAnexo(R.Anexo)
            TA.TerminaContrato(R.Anexo)
            TA.BloqueaAnexo(R.Anexo)
            If R.SaldoFac < 10 Then
                For Each rr As ProduccionDS.GEN_CorreosFasesRow In TCorreosFases.Rows
                    Utilerias.EnviacORREO(rr.Correo, R.AnexoCon, "Terminación de Contrato: " & R.AnexoCon, "Notificaciones@finagil.com.mx")
                Next
            End If
        Next
        'Cancelados por Adelanto a capital
        TA.FillSinSaldoInsoluto(T, Today.AddDays(1).ToString("yyyyMMdd"))
        CorreosFases.Fill(TCorreosFases, "TERMINACONTRATO")
        For Each R As ProduccionDS.AnexosTerminadosRow In T.Rows
            TA.DesbloqueaAnexo(R.Anexo)
            TA.CancelaContrato(R.Anexo)
            TA.BloqueaAnexo(R.Anexo)
            If R.SaldoFac < 10 Then
                For Each rr As ProduccionDS.GEN_CorreosFasesRow In TCorreosFases.Rows
                    Utilerias.EnviacORREO(rr.Correo, R.AnexoCon, "Cancelación de Contrato: " & R.AnexoCon, "Notificaciones@finagil.com.mx")
                Next
            End If
        Next
        TA.QuitaOpciones()
    End Sub

    Public Sub Terminados_Con_Saldo(fec As Date)
        Dim TA As New ProduccionDSTableAdapters.Vw_TerminadosConSaldoTableAdapter
        Dim T As New ProduccionDS.Vw_TerminadosConSaldoDataTable
        CorreosFases.Fill(TCorreosFases, "TERMINACONTRATO")

        TA.Fill(T)
        For Each R As ProduccionDS.Vw_TerminadosConSaldoRow In T.Rows
            TA.TerminaAnexoConSaldo(R.Anexo)
            Utilerias.EnviacORREO("ecacerest@finagil.com.mx", R.Anexo, "Terminación de Contrato con Saldo: " & R.Anexo, "Notificaciones@finagil.com.mx")
        Next

        TA.TermiandosConSaldo_liquidados(T)
        For Each R As ProduccionDS.Vw_TerminadosConSaldoRow In T.Rows
            TA.TerminaAnexo(R.Anexo)
            For Each rr As ProduccionDS.GEN_CorreosFasesRow In TCorreosFases.Rows
                Utilerias.EnviacORREO(rr.Correo, R.Anexo, "Terminación de Contrato (Saldo Pagado): " & R.Anexo, "Notificaciones@finagil.com.mx")
            Next
        Next
        'Termina AV
        TA.TerminaContratosAV_W()
        TA.TerminaContratosConSaldoAV(fec.ToString("yyyyMMdd"))
    End Sub

End Module
