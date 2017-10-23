Module TerminaContratos

    Public Sub Termina_Contratos()
        Dim TA As New ProduccionDSTableAdapters.AnexosTerminadosTableAdapter
        Dim T As New ProduccionDS.AnexosTerminadosDataTable
        Dim ta2 As New WEB_FinagilDSTableAdapters.CorreosTableAdapter
        Dim t2 As New WEB_FinagilDS.CorreosDataTable

        TA.Fill(T, Today.AddDays(10).ToString("yyyyMMdd"))
        ta2.Fill(t2, "TERMINACONTRATO")
        For Each R As ProduccionDS.AnexosTerminadosRow In T.Rows
            TA.DesbloqueaAnexo(R.Anexo)
            TA.TerminaContrato(R.Anexo)
            TA.BloqueaAnexo(R.Anexo)
            For Each rr As WEB_FinagilDS.CorreosRow In t2.Rows
                Utilerias.EnviacORREO(rr.Correo, R.AnexoCon, "Terminación de Contrato: " & R.AnexoCon, "Notificaciones@finagil.com.mx")
            Next
        Next
        TA.QuitaOpciones()
    End Sub

End Module
