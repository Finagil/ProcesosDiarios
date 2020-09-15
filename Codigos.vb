Module Codigos
    Public Sub AltaCodigo(Ruta As String)
        'd_codigo|d_asenta|d_tipo_asenta|D_mnpio|d_estado|d_ciudad|d_CP|c_estado|c_oficina|c_CP|c_tipo_asenta|c_mnpio|id_asenta_cpcons|d_zona|c_cve_ciudad
        Dim ta As New ProduccionDSTableAdapters.CodigosTableAdapter
        Dim t As New ProduccionDS.CodigosDataTable
        Dim sline As String
        Dim Arr As String()
        Dim r As ProduccionDS.CodigosRow
        Try
            Dim Arch As New System.IO.StreamReader(Ruta, System.Text.Encoding.Default)
            Do
                sline = Arch.ReadLine()
                If Not sline Is Nothing Then
                    Arr = sline.Split("|")
                    If IsNumeric(Arr(0)) Then
                        ta.Fill(t, Arr(0), Arr(1), Arr(2))
                        If t.Rows.Count > 0 Then
                            r = t.Rows(0)
                            ta.DeleteCodigo(r.Copos, r.Asentamiento, r.Tipo)
                        End If
                        ta.Insert(Arr(0).Trim, Arr(1).Trim, Arr(2).Trim, Arr(3).Trim, Arr(5).Trim, Arr(4).Trim)
                    End If
                    Console.WriteLine(sline)
                End If
            Loop Until sline Is Nothing
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try




    End Sub
End Module
