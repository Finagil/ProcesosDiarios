Imports System.Data.SqlClient
Imports System.Math
Module CFDI33
    Dim drUdis As DataRowCollection
    Dim nIDSerieA As Decimal = 0
    Dim nIDSerieMXL As Decimal = 0
    Dim cSerie As String = ""
    Dim cSucursal As String = ""
    Dim nTasaIVACliente As Decimal = 0

    Sub FacturarCFDI(FechaProc As Date)
        Dim TaAvisos As New ProduccionDSTableAdapters.AvisosCFDITableAdapter
        Dim TaUdis As New ProduccionDSTableAdapters.TraeUdisTableAdapter
        Dim ProdDS As New ProduccionDS
        'Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection(My.Settings.ProductionCS)
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim daSeries As New SqlDataAdapter(cm1)
        Dim daFacturas As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet()
        Dim dtMovimientos As New DataTable("Movimientos")
        'Dim drMovimiento As DataRow
        'Dim drSaldo As DataRow
        Dim drSerie As DataRow
        Dim strUpdate As String = ""
        Dim strInsert As String = ""
        Dim InstrumentoMonetario As String = ""

        ' Declaración de variables de datos

        Dim cBanco As String = ""
        Dim cCheque As String = ""
        Dim cAnexo As String = ""
        Dim cReferencia As String = ""
        Dim cLetra As String = ""
        Dim cTipar As String = ""
        Dim cTipo As String = ""
        Dim nImporte As Decimal = 0
        Dim nSaldo As Decimal = 0
        Dim nDiasMoratorios As Decimal = 0
        Dim nTasaMoratoria As Decimal = 0
        Dim nMoratorios As Decimal = 0
        Dim nIvaMoratorios As Decimal = 0
        Dim nMontoPago As Decimal = 0
        Dim cFeven As String = ""
        Dim cFepag As String = ""
        Dim cFechaPago As String = ""
        Dim cFechaAplicacion As String = ""
        Dim i As Integer = 0
        Dim nRecibo As Decimal = 0
        Dim Insuficiente As Boolean = False

        cFechaAplicacion = FechaProc.ToString("yyyyMMdd")

        ' Primero creo la tabla Movimientos que contendrá los registros contables de la cobranza

        dtMovimientos.Columns.Add("Anexo", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Letra", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Tipos", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Fepag", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Cve", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Imp", Type.GetType("System.Decimal"))
        dtMovimientos.Columns.Add("Tip", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Catal", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Esp", Type.GetType("System.Decimal"))
        dtMovimientos.Columns.Add("Coa", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Tipmon", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Banco", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Concepto", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Factura", Type.GetType("System.String"))

        ' El siguiente Command trae los consecutivos de cada Serie

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT IDSerieA, IDSerieMXL FROM Llaves"
            .Connection = cnAgil
        End With

        ' Llenar el DataSet lo cual abre y cierra la conexión

        daSeries.Fill(dsAgil, "Series")

        ' Toma el número consecutivo de facturas de pago -que depende de la Serie- y lo incrementa en uno

        drSerie = dsAgil.Tables("Series").Rows(0)
        nIDSerieA = drSerie("IDSerieA")
        nIDSerieMXL = drSerie("IDSerieMXL")

        ' Solo necesito saber el número de elementos que tiene el DataGridView1

        TaAvisos.Fill(ProdDS.AvisosCFDI, cFechaPago)
        TaUdis.Fill(ProdDS.TraeUdis)
        drUdis = ProdDS.TraeUdis.Rows
        For Each r As ProduccionDS.AvisosCFDIRow In ProdDS.AvisosCFDI.Rows

            'cReferencia = DataGridView1.Rows(i).Cells(3).Value
            'InstrumentoMonetario = DataGridView1.Rows(i).Cells(12).Value 'InstrumentoMonetario
            cAnexo = r.Anexo
            'CG.CargaXCliente(CG.SacaCliente(cAnexo))
            Insuficiente = False

            'If DataGridView1.Rows(i).Cells(0).Value = True And CG.Saldo <= 0 Then

            ' Se trata de un depósito seleccionado para su aplicación aunque pudiera tratarse de un 
            ' contrato que adeude más de una letra por lo que debe aplicar el pago en forma
            ' individualizada

            cFechaPago = cFechaAplicacion
            cBanco = ""
            cReferencia = ""
            nImporte = r.SaldoFac
            cCheque = "Facturacion CFDI"
            'If cCheque = "" Then
            '        cCheque = "DR " + cFechaPago
            '        If DataGridView1.Rows(i).Cells(11).Value = True Then
            '            cCheque = "EF " + cFechaPago
            '        End If
            '    Else
            '        cCheque = Mid(cFechaPago, 7, 2) + Mid(cFechaPago, 5, 2) + " " + cCheque
            '        If DataGridView1.Rows(i).Cells(11).Value = True Then
            '            cCheque = "EF " + Mid(cFechaPago, 7, 2) + Mid(cFechaPago, 5, 2)
            '        End If
            '    End If
            'cAnexo = Mid(cReferencia, 1, 5) + Mid(cReferencia, 7, 4)

            'With cm2
            '        .CommandType = CommandType.Text
            '        .CommandText = "SELECT Facturas.Anexo, Letra, Factura, Feven, Fepag, SaldoFac AS Saldo, 0 AS MontoPago, ((Facturas.Tasa + Facturas.Difer) * 2.0) AS TasaMoratoria, Anexos.Tipar, Clientes.Tipo, Clientes.Sucursal, Clientes.TasaIVACliente FROM Facturas " &
            '                   "INNER JOIN Anexos ON Facturas.Anexo = Anexos.Anexo " &
            '                   "INNER JOIN Clientes ON Anexos.Cliente = Clientes.Cliente " &
            '                   "WHERE Facturas.Anexo = '" & cAnexo & "' AND IndPag <> 'P' AND SaldoFac > 0 " &
            '                   "ORDER BY Facturas.Anexo, Letra"
            '        .Connection = cnAgil
            '    End With

            'daFacturas.Fill(dsAgil, "Facturas")

            'strUpdate = "UPDATE Referenciado SET Aplicado = 'S' "
            'strUpdate = strUpdate & "WHERE Referencia = '" & cReferencia & "' AND Fecha = '" & cFechaPago & "' AND Banco = '" & cBanco & "' AND Importe = " & nImporte

            'cnAgil.Open()
            'cm3 = New SqlCommand(strUpdate, cnAgil)
            'cm3.ExecuteNonQuery()
            'cnAgil.Close()

            cBanco = "02" 'bancomer
            'Select Case cBanco
            '    Case "BANAMEX"
            '        cBanco = "04"
            '    Case "BANCOMER"
            '        cBanco = "02"
            '    Case "BANORTE"
            '        cBanco = "10"
            '    Case "HSBC"
            '        cBanco = "03"
            'End Select

            'For Each drSaldo In dsAgil.Tables("Facturas").Rows


            nDiasMoratorios = 0
            nTasaMoratoria = 0
            nMoratorios = 0
            nIvaMoratorios = 0
            cFeven = r.Feven
            cFepag = r.Feven

            ' Traigo la Sucursal y la Tasa de IVA que aplica al cliente a efecto de poder determinar la Serie a utilizar

            cSucursal = r.Sucursal
            nTasaIVACliente = r.TasaIVACliente


            If cSucursal = "04" Or nTasaIVACliente = 11 Then
                cSerie = "MXL"
            Else
                cSerie = "A"
            End If


            'If Trim(cFepag) = "" Then
            '    nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFeven), CTOD(cFechaPago))
            'Else
            '    If cFeven >= cFepag Then
            '        nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFeven), CTOD(cFechaPago))
            '    Else
            '        nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFepag), CTOD(cFechaPago))
            '    End If
            'End If

            'If nDiasMoratorios < 0 Then
            '    nDiasMoratorios = 0
            'End If

            'If cFechaPago = CANCELA_MORA_DIA_FEST(0) And nDiasMoratorios = CDec(CANCELA_MORA_DIA_FEST(2)) Then ' 'Parametrizado en tabla llaves "Fecha;Domiciliacion:dias"
            '    Select Case CANCELA_MORA_DIA_FEST(1)
            '        Case "V"
            '            If DataGridView1.Rows(i).Cells(9).Value = True Then
            '                nDiasMoratorios = 0
            '            End If
            '        Case "F"
            '            If DataGridView1.Rows(i).Cells(9).Value = False Then
            '                nDiasMoratorios = 0
            '            End If
            '        Case Else
            '            nDiasMoratorios = 0
            '    End Select
            'End If


            nMontoPago = r.SaldoFac

            'If nImporte > 0 And nImporte >= (nMoratorios + nIvaMoratorios) Then

            '    nImporte = nImporte - (nMoratorios + nIvaMoratorios)
            '    If nImporte >= nSaldo Then

            '        nMontoPago = nSaldo + nMoratorios + nIvaMoratorios

            '        nImporte = nImporte - nSaldo
            '    Else
            '        nMontoPago = nImporte + nMoratorios + nIvaMoratorios
            '        nImporte = 0
            '    End If
            'Else
            '    If (nMoratorios + nIvaMoratorios) > 0 And nImporte > 0 Then ' si pasa por esta parte es por que el deposito no alcanza para los moratorios y ya no debe continuar con las aplicaciones #ECT 20151029
            '        Insuficiente = True
            '        Exit For
            '    End If
            'End If

            ' La siguiente condición es para evitar que se generen facturas de pago por pagos menores
            ' o iguales a 30 pesos.

            If nMontoPago > 3 Then
                If cSerie = "A" Then
                    nIDSerieA = nIDSerieA + 1
                    nRecibo = nIDSerieA
                ElseIf cSerie = "MXL" Then
                    nIDSerieMXL = nIDSerieMXL + 1
                    nRecibo = nIDSerieMXL
                End If

                cLetra = r.Letra
                Acepagov(cAnexo, cLetra, nMontoPago, nMoratorios, nIvaMoratorios, cBanco, cCheque, dtMovimientos, cFechaAplicacion, cFechaPago, cSerie, nRecibo, InstrumentoMonetario, FechaProc)
            End If

            'Next

            ' Si al terminar el ciclo anterior nImporte fuera mayor que 0, se trata de un saldo a favor del cliente con dos posibilidades:

            ' 1a) Que sea un saldo menor o igual a 30 pesos en cuyo caso se llevará a Otros Productos como abono

            'If nImporte = 0 And nMontoPago > 0 And nMontoPago <= 30 Then

            '        strInsert = "INSERT INTO Historia(Documento, Serie, Numero, Fecha, Anexo, Letra, Banco, Cheque, Balance, Importe, Observa1, InstrumentoMonetario)"
            '        strInsert = strInsert & " VALUES ('"
            '        strInsert = strInsert & "6" & "', '"
            '        strInsert = strInsert & cSerie & "', "
            '        strInsert = strInsert & nRecibo & ", '"
            '        strInsert = strInsert & cFechaAplicacion & "', '"
            '        strInsert = strInsert & cAnexo & "', '"
            '        strInsert = strInsert & cLetra & "', '"
            '        strInsert = strInsert & cBanco & "', '"
            '        strInsert = strInsert & cCheque & "', '"
            '        strInsert = strInsert & "N" & "', '"
            '        strInsert = strInsert & nMontoPago & "',"
            '        strInsert = strInsert & "'OTROS CARGOS', '" & InstrumentoMonetario & "')"
            '        cm4 = New SqlCommand(strInsert, cnAgil)
            '        cnAgil.Open()
            '        cm4.ExecuteNonQuery()
            '        cnAgil.Close()

            '        drMovimiento = dtMovimientos.NewRow()
            '        drMovimiento("Anexo") = cAnexo
            '        drMovimiento("Letra") = cLetra
            '        drMovimiento("Tipos") = "2"
            '        drMovimiento("Fepag") = cFechaAplicacion
            '        drMovimiento("Cve") = "34"
            '        drMovimiento("Imp") = nMontoPago
            '        drMovimiento("Tip") = "S"
            '        drMovimiento("Catal") = "F"
            '        drMovimiento("Esp") = 0
            '        drMovimiento("Coa") = "1"
            '        drMovimiento("Tipmon") = "01"
            '        drMovimiento("Banco") = cBanco
            '        drMovimiento("Concepto") = cCheque
            '        drMovimiento("Factura") = cSerie & nRecibo '#ECT para ligar folios Fiscales
            '        dtMovimientos.Rows.Add(drMovimiento)

            '    End If

            ' 2a. Que sea un saldo mayor a 30 pesos el cual el sistema ya no aplicaría porque tendría que hacerse una aplicación manual

            'dsAgil.Tables.Remove("Facturas")
            If cSerie = "A" And nRecibo <> 0 Then
                strUpdate = "UPDATE Llaves SET IDSerieA = " & nRecibo
            ElseIf cSerie = "MXL" And nRecibo <> 0 Then
                strUpdate = "UPDATE Llaves SET IDSerieMXL = " & nRecibo
            End If
            TaAvisos.FacturarAviso(True, cSerie.Trim, nRecibo, r.Factura, r.Anexo)
            cm1 = New SqlCommand(strUpdate, cnAgil)
            cnAgil.Open()
            cm1.ExecuteNonQuery()
            cnAgil.Close()
        Next

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()
        cm4.Dispose()





    End Sub

End Module
