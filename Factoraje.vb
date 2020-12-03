Module Factoraje

    Sub NotificacionFactorajeFACT_VENC()
        'CANCELACIONES..................................................
        Dim TaWEB As New WEB_FinagilDSTableAdapters.CorreosTableAdapter
        Dim TaNotifi As New WEB_FinagilDSTableAdapters.CancelacionesORDTableAdapter
        Dim Notifi As New WEB_FinagilDS.CancelacionesORDDataTable
        Dim r As WEB_FinagilDS.CancelacionesORDRow
        Dim Grupos As New WEB_FinagilDS.CorreosDataTable
        Dim rr As WEB_FinagilDS.CorreosRow
        Dim Mensaje As String

        TaNotifi.Fill(Notifi)
        If Notifi.Rows.Count > 0 And Date.Now.DayOfWeek = DayOfWeek.Monday Then
            TaWEB.Fill(Grupos, "CANCELACIONES")
            Mensaje = "<body style=""font-family: 'calibri',Garamond, 'Comic Sans'"">"
            Mensaje += "Notificación de Facturas Canceladas (Factoraje) " & Date.Now.ToShortDateString & "<br><br>"
            Mensaje += "<table style=""width:"" border='1'>"
            Mensaje += "<tr><td><b>Factura</b></td><td><b>Cliente</b></td><td><b>Fecha Factura</b></td><td><b>Fecha Cancelación</b></td><td><b>Monto Factura</b></td><td><b>Anticipo</b></td><td><b>Cancelación</b></td></tr>"
            For Each r In Notifi.Rows
                Mensaje += "<tr>"

                Mensaje += "<td>" & r.Factura.Trim & "</td><td>" & r.Nombre.Trim & "</td><td ALIGN=center>" & r.FechaFactura.ToShortDateString &
                    "</td><td ALIGN=center>" & r.FechaCancelacion.ToShortDateString & "</td><td ALIGN=right>" & r.ImporteFactura.ToString("n2") &
                    "</td><td ALIGN=right>" & r.ImporteAnticipo.ToString("n2") & "</td><td ALIGN=right>" & r.Serie.Trim & r.NoCancelacion & "</td>"

                Mensaje += "</tr>"
            Next
            Mensaje += "</Table>"
            TaWEB.Fill(Grupos, "CANCELACIONES")
            For Each rr In Grupos.Rows
                MGlobal.enviacorreo(rr.Correo, Mensaje, "Notificación de Facturas Canceladas (Factoraje): " & Date.Now.ToShortDateString, "Notificaciones@finagil.com.mx")
            Next
            TaNotifi.Enviados()
        End If

        ''EN DESUSO..................................................
        'Dim TaWEB As New WEB_FinagilDSTableAdapters.CorreosTableAdapter
        'Dim TaNotifi As New WEB_FinagilDSTableAdapters.FacturasConSaldoTableAdapter
        'Dim Notifi As New WEB_FinagilDS.FacturasConSaldoDataTable
        'Dim r As WEB_FinagilDS.FacturasConSaldoRow
        'Dim Grupos As New WEB_FinagilDS.CorreosDataTable
        'Dim rr As WEB_FinagilDS.CorreosRow
        'Dim Mensaje As String

        'TaNotifi.Fill(Notifi)
        'If Notifi.Rows.Count > 0 And Date.Now.DayOfWeek <> DayOfWeek.Saturday And Date.Now.DayOfWeek <> DayOfWeek.Sunday Then
        '    TaWEB.Fill(Grupos, "FACT_VENC")
        '    Mensaje = "<body style=""font-family: 'calibri',Garamond, 'Comic Sans'"">"
        '    Mensaje += "Notificación de Facturas Vencidas (Factoraje) " & Date.Now.ToShortDateString & "<br><br>"
        '    Mensaje += "<table style=""width:"" border='1'>"
        '    Mensaje += "<tr><td><b>Factura</b></td><td><b>Cliente</b></td><td><b>Fecha Factura</b></td><td><b>Fecha Vencimiento</b></td><td><b>Monto Factura</b></td><td><b>Saldo Factura</b></td><td><b>Dias Retraso</b></td></tr>"
        '    For Each r In Notifi.Rows
        '        Mensaje += "<tr>"

        '        Mensaje += "<td>" & r.Factura.Trim & "</td><td>" & r.Nombre.Trim & "</td><td ALIGN=center>" & r.FechaFactura.ToShortDateString &
        '            "</td><td ALIGN=center>" & r.FechaVencimiento.ToShortDateString & "</td><td ALIGN=right>" & r.ImporteFactura.ToString("n2") &
        '            "</td><td ALIGN=right>" & r.Saldo.ToString("n2") & "</td><td ALIGN=right>" & r.Dias & "</td>"


        '        Mensaje += "</tr>"
        '    Next
        '    Mensaje += "</Table>"
        '    TaWEB.Fill(Grupos, "FACT_VENC")
        '    For Each rr In Grupos.Rows
        '        MGlobal.enviacorreo(rr.Correo, Mensaje, "Notificación de Facturas Vencidas (Factoraje): " & Date.Now.ToShortDateString, "Notificaciones@finagil.com.mx")
        '    Next
        'End If

    End Sub

End Module
