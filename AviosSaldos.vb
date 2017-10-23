Module AviosSaldos
    Public Sub SaldosAvios()
        Dim ta As New ProduccionDSTableAdapters.SaldosAviosTableAdapter
        Dim t As New ProduccionDS.SaldosAviosDataTable
        Dim R As ProduccionDS.SaldosAviosRow
        Dim tx As New ProduccionDSTableAdapters.AVI_SaldosTMPTableAdapter
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

    End Sub
End Module
