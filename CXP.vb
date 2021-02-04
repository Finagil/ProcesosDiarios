Module CXP
    Dim CorreoB As Boolean = False
    Dim Mensaje, MensajeAux, Asunto As String

    Public Sub Inserta_CXP_MOVS_FACT()
        '1 Pendiente
        '2 VoboMC (usuario Factor libera o cambia importe)
        '3 MesaControl (Pasa pa MC)
        '4 AutorizadoMC
        '5 ErrorEnCuenta (MC, AUTMC)
        '6 Procesado (ya esta en CXP)
        '7 Rechazado (no paso a CXP)
        Dim CorreoMC As String = SacaCorreoFase("MCONTROL_CXP")
        Dim Aux() As String = CorreoMC.Split("<")
        Dim UsuarioMC() As String = Aux(1).Split("@")
        Dim ds As New CXP_DS
        Dim taCliFact As New CXP_DSTableAdapters.CXP_FactorClientesMCTableAdapter
        Dim TaPAg As New CXP_DSTableAdapters.CXP_PagosTesoreriaTableAdapter
        Dim taCuent As New CXP_DSTableAdapters.CXP_CuentasBancariasTableAdapter
        Dim taProv As New CXP_DSTableAdapters.CXP_ProveedoresTableAdapter
        Dim TaPags As New Factor100DSTableAdapters.Vw_PagosFactor100TableAdapter
        Dim tPags As New Factor100DS.Vw_PagosFactor100DataTable
        Dim rCta As CXP_DS.CXP_CuentasBancariasRow
        Dim cFecha As String = Today.ToString("yyyyMMdd")
        TaPags.UpdateMoneda()

        TaPags.Fill(tPags, "Pendiente")
        BorraDatos()
        For Each r As Factor100DS.Vw_PagosFactor100Row In tPags.Rows
            taCuent.Fill(ds.CXP_CuentasBancarias, r.clabe, r.rfc)
            If ds.CXP_CuentasBancarias.Rows.Count <= 0 Then
                MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada.", "ErrorEnCuenta: " & r.NOMBRE & "-" & r.clabe)
                MandaCorreoFase("Factoraje@cmoderna.com", "FactorCXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada.", "ErrorEnCuenta: " & r.NOMBRE & "-" & r.clabe)
                TaPags.UpdateEstatus("ErrorEnCuenta", r.id)
            Else
                taCliFact.Fill(ds.CXP_FactorClientesMC, r.cliente)
                If ds.CXP_FactorClientesMC.Rows.Count > 0 Then
                    TaPags.UpdateEstatus("VoboMC", r.id)
                Else
                    rCta = ds.CXP_CuentasBancarias.Rows(0)
                    Asunto = "Requiere Dispersión de FACTORAJE: " & Date.Now
                    CorreoB = True
                    Mensaje += "Solicitud: " & r.referencia & "<br>"
                    Mensaje += "Beneficiario: " & r.NOMBRE & "<br>"
                    Mensaje += "Importe: " & CDec(r.importe).ToString("n2") & "<br><br>"
                    TaPAg.InsertPago("FAC", r.id, rCta.idCuentas, r.importe, Today.Date, Today.Date, r.moneda, Date.Now, r.referencia, rCta.idProveedor)
                    TaPags.UpdateEstatus("Procesado", r.id)
                End If
            End If
        Next
        If CorreoB = True Then
            MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", Asunto, Mensaje)
            MandaCorreoFase("Factoraje@cmoderna.com", "TESORERIA_CXP", Asunto, Mensaje)
        End If

        BorraDatos()
        TaPags.Fill(tPags, "MesaControl")
        For Each r As Factor100DS.Vw_PagosFactor100Row In tPags.Rows
            taCuent.Fill(ds.CXP_CuentasBancarias, r.clabe, r.rfc)
            If ds.CXP_CuentasBancarias.Rows.Count <= 0 Then
                MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada.", "ErrorEnCuentaMC: " & r.NOMBRE & "-" & r.clabe)
                MandaCorreoFase("Factoraje@cmoderna.com", "FactorCXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada.", "ErrorEnCuentaMC: " & r.NOMBRE & "-" & r.clabe)
                TaPags.UpdateEstatus("ErrorEnCuentaMC", r.id)
            Else
                taCliFact.Fill(ds.CXP_FactorClientesMC, r.cliente)
                If ds.CXP_FactorClientesMC.Rows.Count > 0 Then
                    CorreoB = True
                    Asunto = "Requiere autorizacion de FACTORAJE: " & Date.Now
                    Mensaje += "Solicitud: " & r.referencia & "<br>"
                    Mensaje += "Beneficiario: " & r.NOMBRE & "<br>"
                    Mensaje += "Importe: " & CDec(r.importe).ToString("n2") & "<br>"
                    Mensaje += "<A HREF='https://finagil.com.mx/WEBtasas/5Afdb804-9cXp.aspx?User=" & UsuarioMC(0) & "&ID1=0'>Liga para Autorización.</A><br><br>"
                    TaPags.UpdateEstatus(UsuarioMC(0), r.id)
                Else
                    MensajeAux = "Solicitud: " & r.referencia & "<br>"
                    MensajeAux += "Beneficiario: " & r.NOMBRE & "<br>"
                    MensajeAux += "Importe: " & CDec(r.importe).ToString("n2") & "<br>"
                    MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", "Solicitud en el LIMBO", MensajeAux)
                End If
            End If
        Next
        If CorreoB = True Then
            EnviacORREO(CorreoMC, Mensaje, Asunto, "Factoraje@cmoderna.com")
            MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", Asunto, Mensaje)
        End If
        'TaPags.Fill(tPags, "usuario MCONTROL_CXP") esto se hace manual desde wEB TASAS

        BorraDatos()
        TaPags.Fill(tPags, "AutorizadoMC")
        For Each r As Factor100DS.Vw_PagosFactor100Row In tPags.Rows
            taCuent.Fill(ds.CXP_CuentasBancarias, r.clabe, r.rfc)
            If ds.CXP_CuentasBancarias.Rows.Count <= 0 Then
                MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada. (vobo)", "ErrorEnCuentaMC2: " & r.cliente & "-" & r.clabe)
                MandaCorreoFase("Factoraje@cmoderna.com", "FactorCXP", "Beneficiario sin cuenta bancaria o la cuenta no esta autorizada.(vobo)", "ErrorEnCuentaMC2: " & r.cliente & "-" & r.clabe)
                TaPags.UpdateEstatus("ErrorEnCuentaAutMC", r.id)
            Else
                rCta = ds.CXP_CuentasBancarias.Rows(0)
                CorreoB = True
                Asunto = "Requiere Dispersión de FACTORAJE: " & Date.Now
                Mensaje += "Solicitud: " & r.referencia & "<br>"
                Mensaje += "Beneficiario: " & r.NOMBRE & "<br>"
                Mensaje += "Importe: " & CDec(r.importe).ToString("n2") & "<br><br>"
                TaPAg.InsertPago("FAC", r.id, rCta.idCuentas, r.importe, Today.Date, Today.Date, r.moneda, Date.Now, r.referencia, rCta.idProveedor)
                TaPags.UpdateEstatus("Procesado", r.id)
            End If
        Next
        If CorreoB = True Then
            MandaCorreoFase("Factoraje@cmoderna.com", "SISTEMAS_CXP", Asunto, Mensaje)
            MandaCorreoFase("Factoraje@cmoderna.com", "TESORERIA_CXP", Asunto, Mensaje)
        End If

    End Sub
    Public Function SacaCorreoFase(Fase As String) As String
        SacaCorreoFase = ""
        taFASES.Fill(tFASES, Fase)
        For Each rFASES In tFASES.Rows
            SacaCorreoFase = rFASES.Correo
        Next
        Return SacaCorreoFase
    End Function
    Sub BorraDatos()
        Mensaje = ""
        Asunto = ""
        correoB = False
    End Sub

End Module
