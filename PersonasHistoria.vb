Imports System.Net.Mail
Module PersonasHistoria
    Dim Todo As String = "Fecha" ' Todo, Fecha,
    Dim Fecha As String = "20050101"
    Dim FechaD As Date


    Sub Main()
        If Todo = "Fecha" Then
            Fecha = Date.Now.AddDays(-16).ToString("yyyyMMdd")
        End If
        FechaD = CTOD(Fecha)
        Call GeneraPersonas()
        Call Actualiza_saldos("Todo")
        Call Actualiza_saldos("ConSaldo")
        Call Actualiza_saldos("Actuales")
    End Sub

    Sub GeneraPersonas()
        Dim TaAnexos As New ProduccionDSTableAdapters.AnexosTableAdapter
        Dim TaClientes As New ProduccionDSTableAdapters.ClientesTableAdapter
        Dim TaPersonas As New ProduccionDSTableAdapters.HistoriaPersonasTableAdapter
        Dim Tanexos As New ProduccionDS.AnexosDataTable
        Dim TClientes As New ProduccionDS.ClientesDataTable
        Dim RCli As ProduccionDS.ClientesRow
        Dim taAvales As New ProduccionDSTableAdapters.AvalesPLDTableAdapter
        Dim tAvales As New ProduccionDS.AvalesPLDDataTable
        Dim Cad As String
        Dim cAnexo As String
        Dim Cuantos As Double
        Dim Cont As Double = 1
        Dim FechaCon As Date
        Dim Atraso As String
        TaPersonas.DeleteFecha(FechaD)
        TaAnexos.Fill(Tanexos, Fecha)
        Cuantos = Tanexos.Rows.Count
        For Each rAne As ProduccionDS.AnexosRow In Tanexos.Rows
            TaClientes.Fill(TClientes, rAne.Cliente)
            RCli = TClientes.Rows(0)
            Console.WriteLine("Fase 1 Personas: " & rAne.Anexo & " " & Math.Round((Cont / Cuantos) * 100, 2) & "%")
            Cont += 1
            Cad = Trim(RCli.Descr)
            FechaCon = CTOD(rAne.Fechacon)
            If Right(rAne.Anexo, 1) = "-" Then
                cAnexo = Left(rAne.Anexo, rAne.Anexo.Length - 1)
            Else
                cAnexo = rAne.Anexo
            End If
            If cAnexo = "03173/0001" Then
                cAnexo = "03173/0001"
            End If
            Try


                TaPersonas.Insert(cAnexo, Cad, "Acreditado", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If Trim(RCli.Nomaval1) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomaval1), "Aval 1", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If Trim(RCli.Nomrava1) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrava1), "Rep. Aval 1", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)

                If Trim(RCli.Nomaval2) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomaval2), "Aval 2", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If Trim(RCli.Nomrava2) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrava2), "Rep. Aval 2", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)

                If Trim(RCli.Nomrepr) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrepr), "Representante", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If Trim(RCli.Nomrepr2) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrepr2), "Representante 2", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)

                If Trim(RCli.NomObli) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.NomObli), "Obligado", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If Trim(RCli.NomrObl) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.NomrObl), "Rep. Obligado", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                If RCli.Coac = "C" Then
                    If Trim(RCli.Nomcoac) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomcoac), "Coacreditado", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                    If Trim(RCli.Nomrcoac) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrcoac), "Rep. Coacreditado", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                Else
                    If Trim(RCli.Nomcoac) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomcoac), "Aval", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                    If Trim(RCli.Nomrcoac) <> "" Then TaPersonas.Insert(cAnexo, Trim(RCli.Nomrcoac), "Rep. Aval", FechaCon, rAne.Flcan, rAne.Tipar, 0, Cad, Atraso)
                End If
            Catch ex As Exception
                EnviaError("Ecacerest@Finagil.com.mx", ex.Message, "error de GeneraPersonas " & cAnexo)
            End Try
        Next
        'taAvales.FillByAll(tAvales)
        taAvales.Fill(tAvales, Fecha)
        Cuantos = tAvales.Rows.Count
        Cont = 0
        For Each rr As ProduccionDS.AvalesPLDRow In tAvales.Rows
            Try
                Cont += 1
                Console.WriteLine("Fase 2 Personas: " & rr.Anexo & " " & Math.Round((Cont / Cuantos) * 100, 2) & "%")
                TaPersonas.DeletePersona(rr.Anexo, rr.Persona.Trim, rr.DescripPers.Trim)
                FechaCon = CTOD(rr.Fechacon)
                TaPersonas.Insert(rr.Anexo, rr.Persona.Trim, rr.DescripPers, FechaCon, rr.Flcan, rr.Tipar, 0, rr.Acreditado.Trim, Atraso)
            Catch ex As Exception
                EnviaError("Ecacerest@Finagil.com.mx", ex.Message, "error de GeneraPersonas " & rr.Anexo)
            End Try
        Next


    End Sub

    Sub Actualiza_saldos(Modo As String)
        Dim taEquipo As New ProduccionDSTableAdapters.TablaEquipo1TableAdapter
        Dim taSeg As New ProduccionDSTableAdapters.TablaSeguro1TableAdapter
        Dim tEquipo As New ProduccionDS.TablaEquipo1DataTable
        Dim tSeg As New ProduccionDS.TablaSeguro1DataTable
        Dim TaAnexos As New ProduccionDSTableAdapters.AnexosTableAdapter
        Dim taPersonas As New ProduccionDSTableAdapters.HistoriaPersonasTableAdapter
        Dim Tpersonas As New ProduccionDS.HistoriaPersonasDataTable
        Dim nSaldoEquipo, nInteresEquipo, nCarteraEquipo As Decimal
        Dim nSaldoSeguro, nInteresSeguro, nCarteraSeguro As Decimal
        Dim Cuantos As Double
        Dim Cont As Double = 1
        Dim cAnexo As String
        Dim cCiclo As String
        Dim Atraso As String
        Tpersonas.Clear()
        If Todo = "Todo" And Modo = "Todo" Then
            taPersonas.FillByAllGRP(Tpersonas)
        Else
            If Modo <> "Todo" Then
                If Modo = "ConSaldo" And Todo = "Fecha" Then
                    taPersonas.FillBySaldo(Tpersonas)
                ElseIf Modo = "Actuales" And Todo = "Fecha" Then
                    taPersonas.FillByFecha(Tpersonas, FechaD)
                End If
            End If
        End If

        Cuantos = Tpersonas.Rows.Count
        For Each r As ProduccionDS.HistoriaPersonasRow In Tpersonas.Rows
            Console.WriteLine("Saldos " & Modo & " " & r.Anexo & " " & Math.Round((Cont / Cuantos) * 100, 2) & "%")
            Cont += 1
            cAnexo = Mid(r.Anexo, 1, 5) & Mid(r.Anexo, 7, 4)
            cCiclo = Mid(r.Anexo, 12, 2)
            nSaldoEquipo = 0
            nSaldoSeguro = 0
            If cCiclo <> "" Then
                nSaldoEquipo = TaAnexos.SaldoAvio(cAnexo, cCiclo)
                If TaAnexos.AtrasoAvio(cAnexo, cCiclo) > 0 Then
                    Atraso = "SI"
                Else
                    Atraso = ""
                End If
            Else
                If r.Estatus = "A" Then
                    taEquipo.Fill(tEquipo, cAnexo)
                    TraeSald(tEquipo, Date.Now.ToString("yyyyMMdd"), nSaldoEquipo, nInteresEquipo, nCarteraEquipo, r.TipoCredito)
                    taSeg.Fill(tSeg, cAnexo)
                    TraeSald(tSeg, Date.Now.ToString("yyyyMMdd"), nSaldoSeguro, nInteresSeguro, nCarteraSeguro, r.TipoCredito)
                Else
                    nSaldoEquipo = TaAnexos.SaldoFactura(cAnexo)
                End If
                If TaAnexos.ScalarAtraso(cAnexo) > 0 Then
                    Atraso = "SI"
                Else
                    Atraso = ""
                End If
            End If
            nSaldoEquipo += nSaldoSeguro
            If nSaldoEquipo <= 30 Then nSaldoEquipo = 0

            taPersonas.UpdateSaldo(nSaldoEquipo, Atraso, r.Anexo)
        Next
    End Sub

    Public Sub TraeSald(ByVal drVencimientos As DataTable, ByVal cFeven As String, ByRef nSaldo As Decimal, ByRef nInteres As Decimal, ByRef nCartera As Decimal, Optional ByVal cTipar As String = "")
        Dim drVencimiento As DataRow
        'esta parte trae la opcion a compra no pagada del Arrendamiento puro
        If drVencimientos.Rows.Count > 0 Then
            drVencimiento = drVencimientos(0)
            If cTipar = "P" Then
                Dim Ta As New ProduccionDSTableAdapters.OpcionesTableAdapter
                cTipar = drVencimiento("Anexo")
                If Ta.SacaOpcion(drVencimiento("Anexo")) > 0 Then
                    nSaldo = Ta.SacaOpcion(drVencimiento("Anexo"))
                End If
            End If
        End If

        ' Esta variable datarow contendrá los datos de 1 vencimiento a la vez, de la tabla Edoctav, Edoctas o Edoctao

        For Each drVencimiento In drVencimientos.Rows
            If (drVencimiento("Feven") >= cFeven And drVencimiento("IndRec") = "S") Or drVencimiento("Nufac") = 0 Then
                nSaldo += drVencimiento("Abcap")
                nInteres += drVencimiento("Inter")
                nCartera += drVencimiento("Abcap") + drVencimiento("Inter")
            End If
        Next
        nSaldo = Math.Round(nSaldo, 2)
        nInteres = Math.Round(nInteres, 2)
        nCartera = Math.Round(nCartera, 2)

    End Sub

    Public Function CTOD(ByVal cFecha As String) As Date

        Dim nDia, nMes, nYear As Integer

        nDia = Val(Right(cFecha, 2))
        nMes = Val(Mid(cFecha, 5, 2))
        nYear = Val(Left(cFecha, 4))

        CTOD = DateSerial(nYear, nMes, nDia)

    End Function

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, Asunto) = 0 Then
            Dim Mensage As New MailMessage("InternoBI2008@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient("192.168.110.1", 25)
            Try
                Cliente.Credentials = New System.Net.NetworkCredential("ecacerest", "c4c3r1t0s", "cmoderna")
                Cliente.Send(Mensage)
            Catch ex As Exception
                'ReportError(ex)
            End Try
        Else
            Console.WriteLine("No se ha encontrado la ruta de acceso de la red")
        End If
    End Sub

End Module
