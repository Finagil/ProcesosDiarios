﻿Module AviosSaldos
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
        Dim ta As New ProduccionDSTableAdapters.VwSegVidaTableAdapter
        Dim t As New ProduccionDS.VwSegVidaDataTable
        Dim R As ProduccionDS.VwSegVidaRow

        ta.Fill(t)
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
                Else
                    ta.UpdateSegVida("S", R.SeguroVida, R.Anexo, R.Ciclo)
                End If
            End If
        Next
    End Sub
End Module
