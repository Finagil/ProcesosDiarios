Imports System.Net.Mail
Imports System.IO
Imports System.Data.SqlClient
Module AviosSaldos
    Dim Banamex As String = ""
    Dim Bancomer As String = ""
    Dim Banorte As String = ""
    Dim BBVCIE As String = ""
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
    End Sub

    Sub Aplica_Seguro_Vida()
        Dim ta As New ProduccionDSTableAdapters.VwSegVidaTableAdapter
        Dim t As New ProduccionDS.VwSegVidaDataTable
        Dim R As ProduccionDS.VwSegVidaRow
        Dim rr As ProduccionDS.GEN_CorreosFasesRow
        Dim Para As String = ""
        Dim asunto As String = ""
        Dim Mensaje As String = ""
        Dim De As String = "SEGUROSVIDA@Finagil.com.mx"
        taFASES.Fill(tFASES, "SEGUROSVIDA")
        For Each rr In tFASES.Rows
            Para += (Trim(rr.Correo)) & ";"
        Next
        ta.Fill(t)
        Console.WriteLine("SEGUROSVIDA")
        For Each R In t.Rows
            If R.Tipo = "M" Then
                ta.UpdateSegVida("N", 0, R.Anexo, R.Ciclo)
            Else
                Dim FechaCon As Date = MGlobal.CTOD(R.Fechacon)
                Dim cad As String = R.RFC.Substring(4, 6)
                If CInt(cad.Substring(0, 2)) <= Date.Now.Year - 2000 Then
                    cad = "20" & cad
                Else
                    cad = "19" & cad
                End If
                Dim FechaNac As Date = MGlobal.CTOD(cad)
                Dim Edad As Integer = DateDiff(DateInterval.Year, FechaNac, FechaCon)
                If Edad >= 70 Then
                    ta.UpdateSegVida("N", 0, R.Anexo, R.Ciclo)
                    asunto = "Contrato sin seguro de Vida " & R.AnexoCon
                    Mensaje = "Contrato Sin seguro de Vida por la edad de Cliente: <br>"
                Else
                    ta.UpdateSegVida("S", R.SeguroVida, R.Anexo, R.Ciclo)
                    asunto = "Contrato con seguro de Vida " & R.AnexoCon
                    Mensaje = "Contrato con seguro de Vida por la edad de Cliente: <br>"
                End If
                Mensaje += "Cliente: " & R.Descr & "<br>"
                Mensaje += "Contrato: " & R.AnexoCon & "<br>"
                Mensaje += "Tipo Crédito: " & R.TipoCredito & "<br>"
                Mensaje += "Fecha de Nacimiento: " & FechaNac.ToShortDateString & "<br>"
                Mensaje += "Edad: " & Edad & "<br>"
                MGlobal.EnviacORREO(Para, Mensaje, asunto, De)
            End If
        Next
    End Sub

    Sub AvisoCC()
        Dim ta As New ProduccionDSTableAdapters.Vw_SaldoCCTableAdapter
        Dim t As New ProduccionDS.Vw_SaldoCCDataTable
        Dim r As ProduccionDS.Vw_SaldoCCRow
        Dim Fecha As String = Date.Now.AddMonths(1).AddDays((Date.Now.Day - 1) * -1).ToString("yyyyMMdd")
        Dim FechaD As Date = Date.Now.AddMonths(1).AddDays(Date.Now.Day * -1)
        Dim res As Object

        ta.Fill(t, FechaD.ToString("yyyyMMdd"))
        For Each r In t.Rows
            Console.WriteLine(r.Anexo & "-" & r.Pagare)
            res = Estado_de_Cuenta_Avio(r.Anexo, r.Pagare, 1, "Jobs", Fecha)
            EdoCtaUno(r, FechaD)
        Next
    End Sub

    Public Function Estado_de_Cuenta_Avio(ByVal cAnexo As String, ByVal cCiclo As String, ByVal Proyectado As Integer, ByVal Usuario As String, Fecha As String)
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionString)
        Dim Res As Object
        Dim cm1 As New SqlCommand()
        With cm1
            .CommandType = CommandType.StoredProcedure
            .CommandText = "dbo.EstadoCuentaAvio"
            .CommandTimeout = 50
            .Parameters.AddWithValue("Anexo", cAnexo)
            .Parameters.AddWithValue("Ciclo", cCiclo)
            .Parameters.AddWithValue("Proyectado", Proyectado)
            .Parameters.AddWithValue("usuario", Usuario)
            .Parameters.AddWithValue("Fecha", Fecha)
            .Connection = cnAgil
        End With
        cnAgil.Open()
        Try
            Res = cm1.ExecuteScalar()
        Catch ex As Exception
            MGlobal.EnviacORREO("ecacerest@finagil.com.mx", ex.Message, "Error AvisosCC", "AvisosCC@Finagil.com.mx")
        End Try
        cnAgil.Close()
        cnAgil.Dispose()
        cm1.Dispose()
        Return (Res)
    End Function

    Private Sub EdoCtaUno(ByRef r As ProduccionDS.Vw_SaldoCCRow, Fechad As Date)
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionString)
        Dim cm1 As New SqlCommand()
        Dim daDetalle As New SqlDataAdapter(cm1)
        Dim newrptEdoCtaNew As New rptEdoCtaNew()
        Dim dsAgil As New DataSet()
        Dim Intereses As Decimal

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT DetalleFINAGIL.*, Tipta, Tasas, DiferencialFINAGIL, UltimoCorte, FechaTerminacion, Nombre_Sucursal, tipar, rtrim(concepto)+ ' - ' + rtrim(Factura) as ConceptoX, fondeo, semilla, ampliacion, sinMoratorios, Clientes.Descr FROM DetalleFINAGIL " &
                           "INNER JOIN Avios ON DetalleFINAGIL.Anexo = Avios.Anexo AND DetalleFINAGIL.Ciclo = Avios.Ciclo " &
                           "INNER JOIN Clientes ON Avios.Cliente = Clientes.Cliente " &
                           "INNER JOIN Sucursales ON Clientes.Sucursal = Sucursales.ID_Sucursal " &
                           "WHERE DetalleFINAGIL.Anexo = '" & r.Anexo & "' AND DetalleFINAGIL.Ciclo = '" & r.Pagare & "' " &
                           "ORDER BY Consecutivo"
            .Connection = cnAgil
        End With
        ReFerencia(r.Anexo)
        daDetalle.Fill(dsAgil, "Detalle")
        For Each rR As DataRow In dsAgil.Tables(0).Rows
            Intereses += rR.Item("Intereses")
        Next
        newrptEdoCtaNew.SummaryInfo.ReportTitle = "Saldo al " & Fechad.ToLongDateString
        newrptEdoCtaNew.SummaryInfo.ReportComments = "Cliente : " & r.Clientes.Trim & Space(1) & " Pagare: " & r.Pagare
        newrptEdoCtaNew.SetDataSource(dsAgil)
        newrptEdoCtaNew.SetParameterValue("Refe1", Banamex)
        newrptEdoCtaNew.SetParameterValue("Refe2", Bancomer)
        newrptEdoCtaNew.SetParameterValue("Refe3", BBVCIE)
        newrptEdoCtaNew.SetParameterValue("Refe4", Banorte)
        newrptEdoCtaNew.SetParameterValue("Tipo", "Tipo: Crédito en Cuenta Corriente")
        newrptEdoCtaNew.SetParameterValue("Fondeo", "Tipo de Recursos: " & r.Fondeotit)
        newrptEdoCtaNew.SetParameterValue("Semilla", "INTERES MENSUAL A PAGAR: " & Intereses.ToString("n2"))
        newrptEdoCtaNew.SetParameterValue("Vencimiento", "Vencimiento: " & MGlobal.CTOD(r.FechaTerminacion).ToShortDateString)
        newrptEdoCtaNew.SetParameterValue("Ciclo", "")
        newrptEdoCtaNew.SetParameterValue("Moratorios", 0)
        newrptEdoCtaNew.SetParameterValue("SeguroVida", 0)
        newrptEdoCtaNew.SetParameterValue("Dias", 0)
        newrptEdoCtaNew.SetParameterValue("TasaMora", 0)
        Dim Archi As String = "\AVISOS\AvisoAV" & r.Anexo & "-" & r.Pagare & ".Pdf"
        Dim Archivo As String = My.Settings.RUTA_TMP & Archi
        File.Delete(Archivo)
        newrptEdoCtaNew.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Archivo)
        Dim PARA As String = ""
        Dim DE As String = "AvisosCC@finagil.com.mx"
        taFASES.Fill(tFASES, "AVISOSCC")
        For Each rr In tFASES.Rows
            PARA += (Trim(rr.Correo)) & ";"
        Next

        Dim Asunto As String = "AVISO " & r.Anexo & "-" & r.Pagare & " " & Fechad.ToString("yyyyMMdd") & " FINAGIL, S.A. de C.V. SOFOM, E.N.R."
        Dim Mensaje As String = "Contrato : " & r.AnexoCon & "<br>" &
            "Pagare : " & r.Pagare & "<br>" &
                "FECHA LIMITE DE PAGO : " & Fechad.ToShortDateString & "<br>" &
                "IMPORTE A PAGAR DE INTERES : " & Intereses.ToString("N2") & "<br>" & "<br>" &
                "ESTIMADO CLIENTE : " & "<br>" &
                "Usted podrá consultar sus facturas CFDI en nuestra página de internet www.finagil.com.mx" & "<br>" &
                "Sin más por el momento agradecemos su atención y nos ponemos a su disposición en el teléfono" & "<br>" &
                "01 722 214 5533 ext. 1010 o al 800 727 7100, en caso de cualquier duda o comentario al respecto" & "<br>"
        MGlobal.EnviacORREO(PARA, Mensaje, Asunto, DE, Archi)
        If InStr(r.EMail1, "@") Then MGlobal.EnviacORREO(r.EMail1.Trim, Mensaje, Asunto, DE, Archi)
        If InStr(r.EMail2, "@") Then MGlobal.EnviacORREO(r.EMail2.Trim, Mensaje, Asunto, DE, Archi)
        If InStr(r.Correo, "@") Then MGlobal.EnviacORREO(r.Correo.Trim, Mensaje, Asunto, DE, Archi)

        newrptEdoCtaNew.Dispose()
        cnAgil.Dispose()
        cm1.Dispose()
    End Sub

    Sub ReFerencia(ByVal cAnexo As String)
        'Parte correspondiente a obtener Las cuentas para Depositos Referenciados

        Dim nResultado As Decimal
        Dim nSumaBanamex As Decimal
        Dim nSumaBancomer As Decimal

        Dim cRefBanamex As String
        Dim cRefBanorte As String
        Dim cRefBancomer As String

        cRefBanamex = Mid(cAnexo, 1, 5) + Mid(cAnexo, 7, 3)
        cRefBancomer = Mid(cAnexo, 2, 4) + Mid(cAnexo, 7, 3)
        cRefBanorte = Mid(cAnexo, 2, 4) + Mid(cAnexo, 7, 3)

        nSumaBanamex = 1235
        nSumaBanamex += Val(Mid(cRefBanamex, 1, 1)) * 11
        nSumaBanamex += Val(Mid(cRefBanamex, 2, 1)) * 13
        nSumaBanamex += Val(Mid(cRefBanamex, 3, 1)) * 17
        nSumaBanamex += Val(Mid(cRefBanamex, 4, 1)) * 19
        nSumaBanamex += Val(Mid(cRefBanamex, 5, 1)) * 23
        nSumaBanamex += Val(Mid(cRefBanamex, 6, 1)) * 29
        nSumaBanamex += Val(Mid(cRefBanamex, 7, 1)) * 31
        nSumaBanamex += Val(Mid(cRefBanamex, 8, 1)) * 37

        nResultado = 99 - (nSumaBanamex Mod 97)
        If nResultado > 9 Then
            cRefBanamex += "-" + nResultado.ToString
        Else
            cRefBanamex += "-" + "0" + nResultado.ToString
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

        cRefBancomer += "-" + nResultado.ToString
        cRefBanorte = cRefBancomer

        Banamex = "BANAMEX		Suc. 285 Cuenta 7944154	Referencia: " & cRefBanamex
        Bancomer = "BANCOMER		Convenio 581034			Referencia: " & cRefBancomer
        Banorte = "BANORTE		CEP 36832				Referencia: " & cRefBanorte
        BBVCIE = "BANCOMER  INTERBANCARIO  Convenio CIE 1244159  CIE Interbancario 012914002012441593  Referencia: " & cRefBancomer

    End Sub
End Module
