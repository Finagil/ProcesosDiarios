Imports System.Data.SqlClient
Imports System.Math
Module CFDI33
    Dim drUdis As DataRowCollection
    Dim nIDSerieA As Decimal = 0
    Dim nIDSerieMXL As Decimal = 0
    Dim cSerie As String = ""
    Dim cSucursal As String = ""
    Dim nTasaIVACliente As Decimal = 0

    Sub FacturarCFDI(FechaProc As Date, Tipo As String)
        Dim TaAvisos As New ProduccionDSTableAdapters.AvisosCFDITableAdapter
        Dim TaUdis As New ProduccionDSTableAdapters.TraeUdisTableAdapter
        Dim ProdDS As New ProduccionDS
        'Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection(My.Settings.ConectionStringCFDI)
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
        Dim MetodoPago As String

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
        cFechaPago = cFechaAplicacion
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
        Select Case Tipo.ToUpper
            Case "PREPAGO" ' prepagos antes de su fecha de vencimiento
                TaAvisos.FillByPrepagos(ProdDS.AvisosCFDI, cFechaPago)
            Case "DIA" 'avisos de vencimiento del dia
                TaAvisos.FillporDia(ProdDS.AvisosCFDI, cFechaPago)
            Case "PASADOS" ' avisos generados despues de su vencimiento
                TaAvisos.FillByAnteriores(ProdDS.AvisosCFDI, cFechaPago)
            Case "PENDIENTES"
                TaAvisos.FillHastaFecha(ProdDS.AvisosCFDI, cFechaPago)
        End Select

        'TaAvisos.FillHastaFecha(ProdDS.AvisosCFDI, cFechaPago)
        TaUdis.Fill(ProdDS.TraeUdis)
        drUdis = ProdDS.TraeUdis.Rows
        For Each r As ProduccionDS.AvisosCFDIRow In ProdDS.AvisosCFDI.Rows

            'cReferencia = DataGridView1.Rows(i).Cells(3).Value
            'InstrumentoMonetario = DataGridView1.Rows(i).Cells(12).Value 'InstrumentoMonetario
            Console.WriteLine("Aviso:" & r.Factura)
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
            cBanco = "02" 'bancomer
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

            If r.Tipar <> "B" Then
                nMontoPago = r.ImporteFac * 2
            Else
                nMontoPago = (r.IvaCapital + r.RenPr) * 2
            End If

            If nMontoPago > 3 Then
                If cSerie = "A" Then
                    nIDSerieA = nIDSerieA + 1
                    nRecibo = nIDSerieA
                ElseIf cSerie = "MXL" Then
                    nIDSerieMXL = nIDSerieMXL + 1
                    nRecibo = nIDSerieMXL
                End If
                MetodoPago = "PPD"
                cLetra = r.Letra
                Acepagov(cAnexo, cLetra, nMontoPago, nMoratorios, nIvaMoratorios, cBanco, cCheque, dtMovimientos, cFechaAplicacion, cFechaPago, cSerie, nRecibo, InstrumentoMonetario, FechaProc, MetodoPago)
            End If

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
