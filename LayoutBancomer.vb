Imports System.IO
Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient
Module LayoutBancomer
    Public Function EnviaLayout(ByVal cTipoReporte As String, Dias As Integer) As String

        Dim cnAgil As New SqlConnection(My.Settings.ConnectionStringDOMI)
        Dim cm1 As New SqlCommand
        Dim cm2 As New SqlCommand
        Dim cm3 As New SqlCommand
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daCorreos As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet
        Dim dtDomiciliacion As New DataTable
        Dim drAnexo As DataRow
        Dim drCorreo As DataRow
        Dim drDomiciliacion As DataRow
        Dim ContadorAux1 As Integer = 1

        Dim strUpdate As String = ""

        Dim Servidor As New SmtpClient
        Dim Mensaje As New MailMessage
        Dim Adjunto As Attachment

        Dim ms As New MemoryStream
        Dim writer As StreamWriter

        Dim cAnexo As String = ""
        Dim cBanco As String = ""
        Dim cCuenta As String = ""
        Dim cDescr As String = ""
        Dim cDia As String = ""
        Dim cFechaFinal As String = ""
        Dim cFechaInicial As String = ""
        Dim cFechaFinalEXT As String = ""
        Dim cFechaInicialEXT As String = ""
        Dim cLetra As String = ""
        Dim cLeyenda As String = ""
        Dim cMensaje As String = ""
        Dim cPago As String = ""
        Dim cRefBancomer As String = ""
        Dim cReferencia As String = ""
        Dim cRenglon As String = ""
        Dim cSumaPago As String = ""
        Dim cTipo As String = ""
        Dim cTitular As String = ""
        Dim lProcesar As Boolean = True
        Dim nCount As Integer = 0
        Dim nPago As Decimal = 0
        Dim nResultado As Decimal = 0
        Dim nSaldoFac As Decimal = 0
        Dim nSumaBancomer As Decimal = 0
        Dim nSumaPago As Decimal = 0
        Dim nIDCargoExtra As Integer = 0

        ' Dado que el job correrá todos los días a las 8:00 a.m. debo omitir sábado y domingo del proceso
        Dim Hoy As Date = Today.AddDays(Dias)
        'Hoy = CDate("30/03/2018") 'PARA PRUEBAS

        Dim nDiaSemana As Byte = Hoy.Date.DayOfWeek

        Select Case Hoy' dias festivos
            Case CDate("16/09/2016")
                Exit Function
            Case CDate("25/12/2016")
                Exit Function
            Case CDate("01/01/2017")
                Exit Function
        End Select

        Select Case nDiaSemana
            Case 0                  ' Domingo
                lProcesar = False
            Case 1                  ' Lunes
                lProcesar = True
            Case 2                  ' Martes
                lProcesar = True
            Case 3                  ' Miércoles
                lProcesar = True
            Case 4                  ' Jueves
                lProcesar = True
            Case 5                  ' Viernes
                lProcesar = True
            Case 6                  ' Sábado
                lProcesar = False
        End Select

        If lProcesar = True Then

            dtDomiciliacion.Columns.Add("Contrato", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Letra", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Vencimiento", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("UltimoPago", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Saldo", Type.GetType("System.Decimal"))
            dtDomiciliacion.Columns.Add("Banco", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Tipo", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Cuenta", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Titular", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Name", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("Referencia", Type.GetType("System.String"))
            dtDomiciliacion.Columns.Add("IDCargoExtra", Type.GetType("System.Int32"))

            If cTipoReporte = "B" Then

                Select Case nDiaSemana
                    Case 1 To 4             ' Lunes a Jueves
                        cFechaInicial = Hoy.AddDays(1).ToString("yyyyMMdd")
                        cFechaFinal = Hoy.AddDays(1).ToString("yyyyMMdd")

                        cFechaInicialEXT = Hoy.AddDays(0).ToString("yyyyMMdd") ' hoy 
                        cFechaFinalEXT = Hoy.AddDays(1).ToString("yyyyMMdd")'mañana
                    Case 5                  ' Viernes
                        cFechaInicial = Hoy.AddDays(1).ToString("yyyyMMdd") 'sabado
                        cFechaFinal = Hoy.AddDays(3).ToString("yyyyMMdd") 'domingo y lunes

                        cFechaInicialEXT = Hoy.AddDays(0).ToString("yyyyMMdd") 'hoy
                        cFechaFinalEXT = Hoy.AddDays(3).ToString("yyyyMMdd") ' sabado, domingo y lunes
                End Select


                With cm1
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT SaldoFac, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, Facturas.Letra, Anexos.Autoriza, Facturas.Anexo, Tipo, Facturas.Feven, Facturas.Fepag, 0 AS [id_Cargo_Extra] FROM Facturas " &
                                   "INNER JOIN Clientes ON Facturas.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = Facturas.Anexo " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = Facturas.Anexo " &
                                   "WHERE Feven >= '" & cFechaInicial & "' AND Feven <= '" & cFechaFinal & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE = '' AND Facturas.SaldoFac > 0 " &
                                   "UNION " &
                                   "SELECT SaldoFac, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, Facturas.Letra, Anexos.Autoriza, Facturas.Anexo, Tipo, Facturas.Feven, Facturas.Fepag, 0 AS [id_Cargo_Extra] FROM Facturas " &
                                   "INNER JOIN Clientes ON Facturas.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = Facturas.Anexo " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = Facturas.Anexo " &
                                   "WHERE Feven >= '" & cFechaInicial & "' AND Feven <= '" & cFechaFinal & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE <> '' AND CuentasDomi.Banco = 'BANCOMER' AND Facturas.SaldoFac > 0 " &
                                   "UNION " &
                                   "SELECT PROM_CARGOS_EXTRAS.ImporteTotal, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, '' AS Letra, Anexos.Autoriza, PROM_CARGOS_EXTRAS.Anexo, Tipo, PROM_CARGOS_EXTRAS.FechaCargo, '' AS Fepag, id_Cargo_Extra FROM PROM_CARGOS_EXTRAS " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "INNER JOIN Clientes ON Anexos.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "WHERE FechaCargo >= '" & cFechaInicialEXT & "' AND FechaCargo <= '" & cFechaFinalEXT & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE = '' AND PROM_CARGOS_EXTRAS.Importe > 0 AND PROM_Cargos_Extras.Procesado = 0" &
                                   "UNION " &
                                   "SELECT PROM_CARGOS_EXTRAS.ImporteTotal, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, '' AS Letra, Anexos.Autoriza, PROM_CARGOS_EXTRAS.Anexo, Tipo, PROM_CARGOS_EXTRAS.FechaCargo, '' AS Fepag, id_Cargo_Extra FROM PROM_CARGOS_EXTRAS " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "INNER JOIN Clientes ON Anexos.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "WHERE FechaCargo >= '" & cFechaInicialEXT & "' AND FechaCargo <= '" & cFechaFinalEXT & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE <> '' AND CuentasDomi.Banco = 'BANCOMER' AND PROM_CARGOS_EXTRAS.Importe > 0 AND PROM_Cargos_Extras.Procesado = 0"
                    .Connection = cnAgil

                End With

            ElseIf cTipoReporte = "O" Then

                Select Case nDiaSemana
                    Case 2 To 5             ' martes a Viernes
                        cFechaInicial = Hoy.AddDays(0).ToString("yyyyMMdd")
                        cFechaFinal = Hoy.AddDays(0).ToString("yyyyMMdd")
                    Case 1                  ' lunes
                        cFechaInicial = Hoy.AddDays(-2).ToString("yyyyMMdd")
                        cFechaFinal = Hoy.AddDays(0).ToString("yyyyMMdd")
                End Select



                With cm1
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT SaldoFac, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, Facturas.Letra, Anexos.Autoriza, Facturas.Anexo, Tipo, Facturas.Feven, Facturas.Fepag, 0 AS [id_Cargo_Extra] FROM Facturas " &
                                   "INNER JOIN Clientes ON Facturas.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = Facturas.Anexo " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = Facturas.Anexo " &
                                   "WHERE Feven >= '" & cFechaInicial & "' AND Feven <= '" & cFechaFinal & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE <> '' AND CuentasDomi.Banco <> 'BANCOMER' AND Facturas.SaldoFac > 0 " &
                                   "UNION " &
                                   "SELECT PROM_CARGOS_EXTRAS.ImporteTotal, Descr, CuentasDomi.Banco, CuentasDomi.CuentaCLABE, CuentasDomi.NumTarjeta, CuentasDomi.CuentaEJE, CuentasDomi.TitularCta, Referencia, rtrim(letra) AS Letra, Anexos.Autoriza, PROM_CARGOS_EXTRAS.Anexo, Tipo, PROM_CARGOS_EXTRAS.FechaCargo, '' AS Fepag, id_Cargo_Extra FROM PROM_CARGOS_EXTRAS " &
                                   "INNER JOIN Anexos ON Anexos.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "INNER JOIN Clientes ON Anexos.Cliente = Clientes.Cliente " &
                                   "INNER JOIN CuentasDomi ON CuentasDomi.Anexo = PROM_CARGOS_EXTRAS.Anexo " &
                                   "WHERE FechaCargo >= '" & cFechaInicial & "' AND FechaCargo <= '" & cFechaFinal & "' AND Anexos.Autoriza = 'S' AND CuentasDomi.CuentaCLABE <> '' AND CuentasDomi.Banco <> 'BANCOMER' AND PROM_CARGOS_EXTRAS.Importe > 0 AND PROM_Cargos_Extras.Procesado = 0"
                    .Connection = cnAgil
                End With

            End If

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "SELECT Correo FROM GEN_CorreosFases " &
                               "WHERE (Fase = 'DOMICILIACION') " &
                               "ORDER BY id_correo"
                .Connection = cnAgil
            End With

            ' Llenar el DataSet a través del DataAdapter lo que abre y cierra la conexión

            daAnexos.Fill(dsAgil, "Pagos")
            daCorreos.Fill(dsAgil, "Correos")
            'Dim ta As New ProduccionDSTableAdapters.AnexosTableAdapter
            Dim Pesos As Decimal
            Dim Particion As Integer = 0
            If dsAgil.Tables("Pagos").Rows.Count > 0 Then

                For Each drAnexo In dsAgil.Tables("Pagos").Rows ' hace vario correos por montos mayores a 
                    'If ta.EsRimo(drAnexo("Anexo")) <= 0 Then
                    '    Continue For
                    'End If
                    Particion = 1
                    nSaldoFac = drAnexo("SaldoFac")
                    Pesos = 50000
                    While nSaldoFac > Pesos
                        drDomiciliacion = dtDomiciliacion.NewRow()
                        drDomiciliacion("Contrato") = drAnexo("Anexo")
                        drDomiciliacion("Letra") = Chr(64 + Particion) & Mid(drAnexo("Letra"), 2, 2)
                        If Trim(drAnexo("Feven")) <> "" Then
                            drDomiciliacion("Vencimiento") = drAnexo("Feven")
                        Else
                            drDomiciliacion("Vencimiento") = "        "
                        End If
                        If Trim(drAnexo("Fepag")) <> "" Then
                            drDomiciliacion("UltimoPago") = drAnexo("Fepag")
                        Else
                            drDomiciliacion("UltimoPago") = "        "
                        End If
                        drDomiciliacion("Saldo") = Pesos
                        drDomiciliacion("Banco") = drAnexo("Banco")
                        drDomiciliacion("Tipo") = drAnexo("Tipo")
                        If Trim(drAnexo("CuentaCLABE")) <> "" Then
                            drDomiciliacion("Cuenta") = drAnexo("CuentaCLABE")
                        ElseIf Trim(drAnexo("NumTarjeta")) <> "" Then
                            drDomiciliacion("Cuenta") = drAnexo("NumTarjeta")
                        Else
                            drDomiciliacion("Cuenta") = drAnexo("CuentaEJE")
                        End If
                        drDomiciliacion("Titular") = drAnexo("TitularCta")
                        drDomiciliacion("Name") = drAnexo("Descr")
                        drDomiciliacion("Referencia") = drAnexo("Referencia")
                        drDomiciliacion("IDCargoExtra") = drAnexo("id_Cargo_Extra")
                        dtDomiciliacion.Rows.Add(drDomiciliacion)

                        nSaldoFac = nSaldoFac - Pesos
                        Pesos -= 1
                        Particion += 1
                    End While

                    If nSaldoFac > 0 Then

                        drDomiciliacion = dtDomiciliacion.NewRow()
                        drDomiciliacion("Contrato") = drAnexo("Anexo")
                        If Particion = 1 Then
                            drDomiciliacion("Letra") = drAnexo("Letra")
                        Else
                            drDomiciliacion("Letra") = Chr(64 + Particion) & Mid(drAnexo("Letra"), 2, 2)
                        End If

                        If Trim(drAnexo("Feven")) <> "" Then
                            drDomiciliacion("Vencimiento") = drAnexo("Feven")
                        Else
                            drDomiciliacion("Vencimiento") = "        "
                        End If
                        If Trim(drAnexo("Fepag")) <> "" Then
                            drDomiciliacion("UltimoPago") = drAnexo("Fepag")
                        Else
                            drDomiciliacion("UltimoPago") = "        "
                        End If
                        drDomiciliacion("Saldo") = nSaldoFac
                        drDomiciliacion("Banco") = drAnexo("Banco")
                        drDomiciliacion("Tipo") = drAnexo("Tipo")
                        If Trim(drAnexo("CuentaCLABE")) <> "" Then
                            drDomiciliacion("Cuenta") = drAnexo("CuentaCLABE")
                        ElseIf Trim(drAnexo("NumTarjeta")) <> "" Then
                            drDomiciliacion("Cuenta") = drAnexo("NumTarjeta")
                        Else
                            drDomiciliacion("Cuenta") = drAnexo("CuentaEJE")
                        End If
                        drDomiciliacion("Titular") = drAnexo("TitularCta")
                        drDomiciliacion("Name") = drAnexo("Descr")
                        drDomiciliacion("Referencia") = drAnexo("Referencia")
                        drDomiciliacion("IDCargoExtra") = drAnexo("id_Cargo_Extra")
                        dtDomiciliacion.Rows.Add(drDomiciliacion)
                    End If
                Next

                nCount = 1
                If cTipoReporte = "B" Then
                    writer = New StreamWriter("c:\files\Pagos_BANCOMER_" & Hoy.ToString("ddMMyyyy") & "_" & ContadorAux1 & ".txt")
                ElseIf cTipoReporte = "O" Then
                    writer = New StreamWriter("c:\files\Pagos_OTROS_BANCOS_" & Hoy.ToString("ddMMyyyy") & "_" & ContadorAux1 & ".txt")
                End If

                For Each drAnexo In dtDomiciliacion.Rows

                    If nSumaPago >= 400000 Then
                        writer.Close()
                        nCount = 1
                        ContadorAux1 += 1
                        nSumaPago = 0
                        If cTipoReporte = "B" Then
                            writer = New StreamWriter("c:\files\Pagos_BANCOMER_" & Hoy.ToString("ddMMyyyy") & "_" & ContadorAux1 & ".txt")
                        ElseIf cTipoReporte = "O" Then
                            writer = New StreamWriter("c:\files\Pagos_OTROS_BANCOS_" & Hoy.ToString("ddMMyyyy") & "_" & ContadorAux1 & ".txt")
                        End If
                    End If

                    cAnexo = drAnexo("Contrato")
                    If Len(cAnexo) > 9 Then
                        cAnexo = Mid(cAnexo, 1, 5) & Mid(cAnexo, 7, 4)
                    End If
                    cLetra = drAnexo("Letra")

                    cReferencia = drAnexo("Referencia")

                    If cReferencia = "C" Then
                        cRefBancomer = "90" + Mid(cAnexo, 1, 5)
                    Else
                        cRefBancomer = Mid(cAnexo, 2, 4) + Mid(cAnexo, 7, 3)
                    End If

                    nSumaBancomer = 0
                    nResultado = Val(Mid(cRefBancomer, 1, 1)) * 2
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 2, 1)) * 1
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 3, 1)) * 2
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 4, 1)) * 1
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 5, 1)) * 2
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 6, 1)) * 1
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If
                    nResultado = Val(Mid(cRefBancomer, 7, 1)) * 2
                    If nResultado > 9 Then
                        nSumaBancomer += Val(Mid(nResultado.ToString, 1, 1)) + Val(Mid(nResultado.ToString, 2, 1))
                    Else
                        nSumaBancomer += nResultado
                    End If

                    If nSumaBancomer > 60 Then
                        nResultado = 70 - nSumaBancomer
                    ElseIf nSumaBancomer > 50 Then
                        nResultado = 60 - nSumaBancomer
                    ElseIf nSumaBancomer > 40 Then
                        nResultado = 50 - nSumaBancomer
                    ElseIf nSumaBancomer > 30 Then
                        nResultado = 40 - nSumaBancomer
                    ElseIf nSumaBancomer > 20 Then
                        nResultado = 30 - nSumaBancomer
                    ElseIf nSumaBancomer > 10 Then
                        nResultado = 20 - nSumaBancomer
                    ElseIf nSumaBancomer > 0 Then
                        nResultado = 10 - nSumaBancomer
                    Else
                        nResultado = 0
                    End If

                    cRefBancomer += nResultado.ToString
                    If Trim(cLetra) <> "" Then
                        cRefBancomer = "PAGO " & cLetra & " DEL CONTRATO " & cRefBancomer
                    Else
                        cLetra = Utilerias.Stuff(nCount.ToString, "I", "0", 3)
                        cRefBancomer = "PAGO " & cLetra & " EXT " & Today.ToString("yyyyMMdd") & " " & cRefBancomer
                    End If

                    nPago = drAnexo("Saldo")

                    cDescr = Mid(Trim(drAnexo("Name")), 1, 40)
                    cTitular = Mid(Trim(drAnexo("Titular")), 1, 40)

                    cBanco = Trim(drAnexo("Banco"))
                    cCuenta = Trim(drAnexo("Cuenta"))

                    If cBanco = "BANCOMER" Then
                        cBanco = "012"
                        cLeyenda = "CARGO DOMICILIADO A BANCO BANCOMER"
                    Else
                        Select Case cBanco
                            Case "BANAMEX"
                                cBanco = "002"
                                cLeyenda = "CARGO DOMICILIADO A BANCO BANAMEX"
                            Case "SANTANDER"
                                cBanco = "014"
                                cLeyenda = "CARGO DOMICILIADO A BANCO SANTANDER"
                            Case "BANJERCITO"
                                cBanco = "019"
                                cLeyenda = "CARGO DOMICILIADO A BANCO BANJERCITO"
                            Case "HSBC"
                                cBanco = "021"
                                cLeyenda = "CARGO DOMICILIADO A BANCO HSBC"
                            Case "BANCO DEL BAJIO"
                                cBanco = "030"
                                cLeyenda = "CARGO DOMICILIADO A BANCO DEL BAJIO"
                            Case "IXE"
                                cBanco = "032"
                                cLeyenda = "CARGO DOMICILIADO A BANCO IXE"
                            Case "INBURSA"
                                cBanco = "036"
                                cLeyenda = "CARGO DOMICILIADO A BANCO INBURSA"
                            Case "INTERACCIONES"
                                cBanco = "037"
                                cLeyenda = "CARGO DOMICILIADO A BANCO INTERACCIONES"
                            Case "BANCA MIFEL"
                                cBanco = "042"
                                cLeyenda = "CARGO DOMICILIADO A BANCO MIFEL"
                            Case "SCOTIABANK"
                                cBanco = "044"
                                cLeyenda = "CARGO DOMICILIADO A BANCO SCOTIABANK"
                            Case "BANORTE"
                                cBanco = "072"
                                cLeyenda = "CARGO DOMICILIADO A BANCO BANORTE"
                            Case "AZTECA"
                                cBanco = "127"
                                cLeyenda = "CARGO DOMICILIADO A BANCO AZTECA"
                            Case "AHORRO"
                                cBanco = "131"
                                cLeyenda = "CARGO DOMICILIADO A BANCO AHORRO"
                            Case "BANCOPPEL"
                                cBanco = "137"
                                cLeyenda = "CARGO DOMICILIADO A BANCO BANCOPPEL"
                        End Select
                    End If

                    If Len(cCuenta) = 18 Then
                        cTipo = "40"
                    ElseIf Len(cCuenta) = 16 Then
                        cTipo = "03"
                    ElseIf Len(cCuenta) = 10 Then
                        cTipo = "01"
                    End If

                    cDia = Mid(Utilerias.DTOC(Today), 7, 2) & Mid(Utilerias.DTOC(Today), 5, 2)

                    If nCount = 1 Then
                        cRenglon = "01000000130012E2" & Mid(cDia, 1, 2) & Utilerias.Stuff(nCount.ToString, "I", "0", 5) & Utilerias.DTOC(Today) & "0100                         " & "FINAGIL SA DE CV SOFOM ENR              " & "FIN 940905AX7     " & Space(182)
                        writer.WriteLine(cRenglon)
                    End If
                    nCount += 1

                    ' En este segmento la intención es transformar el importe (número) en un string que no lleve el punto decimal pero sí los decimales
                    ' aunque estos sean 00.

                    cPago = Int(nPago).ToString

                    If nPago <> Int(nPago) Then
                        ' Se trata de un pago con centavos por lo que hay que multiplicar los centavos por 100 para convertirlos en un entero
                        ' Por ejemplo:
                        ' Si el residual fuera 0.2 al multiplicarlo por cien tendríamos 20
                        ' Si el residual fuera 0.23 al multiplicarlo por cien tendríamos 23
                        ' Si el residual fuera 0.07 al multiplicarlo por cien tendríamos 7 (tendríamos que anteponerle un cero)
                        ' Este nuevo valor lo convertimos a string y lo concatenamos al string de la parte entera 
                        If nPago < 1 Then
                            cPago = CInt(nPago * 100).ToString
                        ElseIf Math.Round(nPago Mod Int(nPago), 2) * 100 < 10 Then
                            cPago = cPago & "0" & Int(Math.Round(nPago Mod Int(nPago), 2) * 100).ToString
                        Else
                            cPago = cPago & Int(Math.Round(nPago Mod Int(nPago), 2) * 100).ToString
                        End If
                    Else
                        ' Se trata de un pago sin centavos
                        cPago = cPago & "00"
                    End If

                    cRenglon = "02" & Utilerias.Stuff(nCount.ToString, "I", "0", 7) & "3001" & Utilerias.Stuff(cPago, "I", "0", 15) & Utilerias.DTOC(Today) & Space(24) & "51" & Utilerias.DTOC(Today) & cBanco
                    cRenglon = cRenglon & cTipo & Utilerias.Stuff(cCuenta, "I", "0", 20) & Utilerias.Stuff(Trim(cTitular), "D", " ", 40) & Utilerias.Stuff(cRefBancomer, "D", " ", 40) & Utilerias.Stuff(cDescr, "D", " ", 40)
                    cRenglon = cRenglon & "000000000000000" & Utilerias.Stuff((nCount - 1).ToString, "I", "0", 7) & Utilerias.Stuff(cLeyenda, "D", " ", 40) & "00" & Space(21)

                    cRenglon = cRenglon.Replace("Ñ", Chr(78))
                    cRenglon = cRenglon.Replace("ñ", Chr(110))
                    writer.WriteLine(cRenglon)
                    nSumaPago += nPago

                    ' Si se trata de un Cargo Extra tengo que ir a la tabla PRO_CARGOS_EXTRAS y marcarlo como procesado

                    If drAnexo("IDCargoExtra") <> 0 Then
                        strUpdate = "UPDATE PROM_CARGOS_EXTRAS SET Procesado = 1 WHERE id_Cargo_Extra = " & drAnexo("IDCargoExtra")
                        cm3 = New SqlCommand(strUpdate, cnAgil)
                        cnAgil.Open()
                        cm3.ExecuteNonQuery()
                        cnAgil.Close()
                    End If

                    If nSumaPago >= 400000 Then
                        nCount += 1
                        cSumaPago = Int(nSumaPago).ToString
                        If nSumaPago <> Int(nSumaPago) Then
                            If Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100 < 10 Then
                                cSumaPago = cSumaPago & "0" & Int(Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100).ToString
                            Else
                                cSumaPago = cSumaPago & Int(Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100).ToString
                            End If
                        Else
                            ' Se trata de un pago sin centavos
                            cSumaPago = cSumaPago & "00"
                        End If

                        cRenglon = "09" & Utilerias.Stuff(nCount.ToString, "I", "0", 7) & "30" & Mid(cDia, 1, 2) & "00001" & Utilerias.Stuff((nCount - 2).ToString, "I", "0", 7) & Utilerias.Stuff(cSumaPago, "I", "0", 18) & Space(17) & Space(240)
                        writer.WriteLine(cRenglon)
                    End If
                Next

                nCount += 1

                ' Hay que hacer la misma validación para convertir la suma de los pagos en string
                cSumaPago = Int(nSumaPago).ToString
                If nSumaPago <> Int(nSumaPago) Then
                    If Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100 < 10 Then
                        cSumaPago = cSumaPago & "0" & Int(Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100).ToString
                    Else
                        cSumaPago = cSumaPago & Int(Math.Round(nSumaPago Mod Int(nSumaPago), 2) * 100).ToString
                    End If
                Else
                    ' Se trata de un pago sin centavos
                    cSumaPago = cSumaPago & "00"
                End If

                cRenglon = "09" & Utilerias.Stuff(nCount.ToString, "I", "0", 7) & "30" & Mid(cDia, 1, 2) & "00001" & Utilerias.Stuff((nCount - 2).ToString, "I", "0", 7) & Utilerias.Stuff(cSumaPago, "I", "0", 18) & Space(17) & Space(240)
                writer.WriteLine(cRenglon)
                writer.Close()
                Try
                    ms.Position = 0
                    Servidor.Host = "smtp01.cmoderna.com"
                    Servidor.Port = "26"
                    For Each drCorreo In dsAgil.Tables("Correos").Rows
                        Mensaje.To.Add(Trim(drCorreo("Correo")))
                    Next
                    'Mensaje.To.Add("ecacerest@finagil.com.mx")
                    Mensaje.From = New MailAddress("Domiciliacion@Finagil.com.mx", "FINAGIL envíos automáticos")
                    If cTipoReporte = "B" Then
                        Mensaje.Subject = "Layout BANCOMER"
                        For x = 1 To ContadorAux1
                            Adjunto = New Attachment("c:\files\Pagos_BANCOMER_" & Hoy.ToString("ddMMyyyy") & "_" & x & ".txt", "text/csv")
                            Mensaje.Attachments.Add(Adjunto)
                        Next
                        'Adjunto = New Attachment(ms, "Pagos BANCOMER_" & Hoy.ToString("ddMMyyyy") & ".txt", "text/csv")
                    ElseIf cTipoReporte = "O" Then
                        Mensaje.Subject = "Layout OTROS BANCOS"
                        For x = 1 To ContadorAux1
                            Adjunto = New Attachment("c:\files\Pagos_OTROS_BANCOS_" & Hoy.ToString("ddMMyyyy") & "_" & x & ".txt", "text/csv")
                            Mensaje.Attachments.Add(Adjunto)
                        Next
                        'Adjunto = New Attachment(ms, "Pagos_OTROS BANCOS_" & Hoy.ToString("ddMMyyyy") & ".txt", "text/csv")
                    End If
                    'Mensaje.Attachments.Add(Adjunto)
                    Servidor.Send(Mensaje)
                    cMensaje = "Generación y envío exitosos"
                Catch ex As Exception
                    cMensaje = ex.Message
                End Try

                writer.Close()
                writer.Dispose()
                ms.Dispose()
                Mensaje.Dispose()
                If Not IsNothing(Adjunto) Then Adjunto.Dispose()
            Else
                Try
                    ms.Position = 0
                    Servidor.Host = "smtp01.cmoderna.com"
                    Servidor.Port = "26"
                    For Each drCorreo In dsAgil.Tables("Correos").Rows
                        Mensaje.To.Add(Trim(drCorreo("Correo")))
                    Next
                    Mensaje.From = New MailAddress("Domiciliacion@Finagil.com.mx", "FINAGIL envíos automáticos (SIN DATOS)")
                    If cTipoReporte = "B" Then
                        Mensaje.Subject = "SIN DATOS - Layout BANCOMER"
                    ElseIf cTipoReporte = "O" Then
                        Mensaje.Subject = "SIN DATOS - Layout OTROS BANCOS"
                    End If
                    Servidor.Send(Mensaje)
                    cMensaje = "Generación y envío exitosos"
                Catch ex As Exception
                    cMensaje = ex.Message
                End Try
            End If

        End If

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()

        Return cMensaje

    End Function

End Module
