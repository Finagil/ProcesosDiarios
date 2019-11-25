Imports System.Data.SqlClient
Module GeneraHistorial
    Sub HistariaConcetrada()
        Dim DSanexos As New DataSet
        Dim Tahisto As New ProduccionDSTableAdapters.ResumenHistoriaTableAdapter
        Dim Historial As New ProduccionDS.ResumenHistoriaDataTable
        Dim r As ProduccionDS.ResumenHistoriaRow
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionString)
        Dim cm0 As New SqlCommand()
        Dim daContratos As New SqlDataAdapter(cm0)
        Dim drContrato As DataRow
        Dim drFactura As DataRow
        Dim drPago As DataRow
        Dim cAnexo As String

        cm0.Connection = cnAgil
        cnAgil.Open()
        cm0.CommandText = "truncate table Temporal1$;"
        cm0.ExecuteNonQuery()

        cm0.CommandType = CommandType.Text
        cm0.CommandText = "select anexo, CLIENTE from anexos where fechacon >= '20090101'"
        daContratos.Fill(DSanexos, "Contratos")
        If DSanexos.Tables("Contratos").Rows.Count < 1 Then
            cnAgil.Close()
            Exit Sub
        End If

        cnAgil.Close()
        Dim XX As Integer = 0
        Dim nNumAtrasos As Double
        Dim nPromDias As Double
        Dim nPromMonto As Double
        Dim nMaxDias As Double
        Dim nMaxMonto As Double
        Dim nNumAdelantos As Double
        Dim nNumPenalizaciones As Double
        Dim nMontoAdelantos As Double
        Dim nMontoPenalizaciones As Double
        Dim Ndias As Double
        Dim nMora As Double
        Dim nNumAmort As Integer

        For Each drContrato In DSanexos.Tables("Contratos").Rows
            cAnexo = drContrato.Item("Anexo")
            nNumAtrasos = 0
            nPromDias = 0
            nPromMonto = 0
            nMaxDias = 0
            nMaxMonto = 0
            nNumAdelantos = 0
            nNumPenalizaciones = 0
            nMontoAdelantos = 0
            nMontoPenalizaciones = 0
            Ndias = 0
            nMora = 0
            nNumAmort = 0

            Dim DSagil As New DataSet
            Dim cm1 As New SqlCommand()
            Dim cm2 As New SqlCommand()
            Dim daHistoria As New SqlDataAdapter(cm1)
            Dim daFacturas As New SqlDataAdapter(cm2)

            cm1.CommandType = CommandType.StoredProcedure
            cm1.CommandText = "Historia1"
            cm1.Connection = cnAgil
            cm1.Parameters.Add("@Anexo", SqlDbType.NVarChar)
            cm1.Parameters(0).Value = cAnexo

            cm2.CommandType = CommandType.StoredProcedure
            cm2.CommandText = "Historia2"
            cm2.Connection = cnAgil
            cm2.Parameters.Add("@Anexo", SqlDbType.NVarChar)
            cm2.Parameters(0).Value = cAnexo
            cm2.Parameters.Add("@Fecha", SqlDbType.NVarChar)
            cm2.Parameters(1).Value = Date.Now.ToString("yyyyMMdd")

            cnAgil.Open()
            daHistoria.Fill(DSagil, "Historia")
            daFacturas.Fill(DSagil, "Facturas")
            daHistoria.Dispose()
            daFacturas.Dispose()

            cm1.Dispose()
            cm2.Dispose()

            For Each drFactura In DSagil.Tables("Facturas").Rows
                Ndias = DateDiff(DateInterval.Day, CTOD(drFactura("Feven")), CTOD(drFactura("Fepag")))
                If Ndias > 0 Then
                    cm0.CommandText = "SELECT isnull(sum(Importe),0) as mora FROM Historia WHERE Observa1 = 'MORATORIOS' and Anexo = '" & cAnexo & "' and letra = '" & drFactura("Letra") & "'"
                    nMora = cm0.ExecuteScalar
                    If nMora > 0 Then
                        nNumAtrasos += 1
                        nPromDias += Ndias
                        If nMaxMonto <= drFactura("ImporteFac") Then
                            nMaxMonto = drFactura("ImporteFac")
                        End If
                        nPromMonto += drFactura("ImporteFac")
                        If nMaxDias <= Ndias Then
                            nMaxDias = Ndias
                        End If
                    End If
                End If
            Next
            cm0.CommandText = "SELECT Count(letra) as Amort FROM EdoctaV WHERE Anexo = '" & cAnexo & "'"
            nNumAmort = cm0.ExecuteScalar
            cnAgil.Close()

            For Each drPago In DSagil.Tables("Historia").Rows
                If Trim(drPago.Item("Observa1")) = "COMISION POR PREPAGO" Or Trim(drPago.Item("Observa1")) = "COMISION POR ADELANTO" Then
                    If drPago.Item("Importe") > 0 Then
                        nNumPenalizaciones += 1
                        nMontoPenalizaciones += drPago.Item("Importe")
                    End If
                ElseIf InStr(drPago.Item("Observa1"), "ADELANTO") > 0 Then
                    If drPago.Item("Importe") > 0 Then
                        nNumAdelantos += 1
                        nMontoAdelantos += drPago.Item("Importe")
                    End If
                End If
            Next
            r = Historial.NewRow
            r.Anexo = cAnexo
            If r.Anexo = "034900001" Then
                nNumAtrasos -= 2
                nMaxDias = 0
                nMaxMonto = 0
            End If
            r.Cliente = drContrato.Item("Cliente")
            r.Atrasos = nNumAtrasos
            If nNumAtrasos > 0 Then
                nPromDias = nPromDias / nNumAtrasos
                nPromMonto = nPromMonto / nNumAtrasos
            Else
                nPromDias = 0
                nPromMonto = 0
            End If
            r.DiasProm = nPromDias
            r.MontoProm = nPromMonto
            r.DiasMax = nMaxDias
            r.MontoMax = nMaxMonto

            r.Adelantos = nNumAdelantos
            r.MontoAdelanto = nMontoAdelantos
            r.Penalizaciones = nNumPenalizaciones
            r.MontoPena = nMontoPenalizaciones
            r.Amortizaciones = nNumAmort
            Historial.AddResumenHistoriaRow(r)
            DSagil.Dispose()
            xx += 1
            If XX >= 1000 Then
                XX = 1
                Historial.GetChanges()
                Tahisto.Update(Historial)
            End If
        Next
        Historial.GetChanges()
        Tahisto.Update(Historial)

    End Sub
End Module
