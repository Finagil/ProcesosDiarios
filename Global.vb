Module MGlobal
    Public taFASES As New ProduccionDSTableAdapters.GEN_CorreosFasesTableAdapter
    Public tFASES As New ProduccionDS.GEN_CorreosFasesDataTable
    Public rFASES As ProduccionDS.GEN_CorreosFasesRow
    Public taCorreos As New ProduccionDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter

    Public Sub EnviacORREO(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String, de As String, Optional Archivo As String = "")
        taCorreos.Insert(de, Para, Asunto, Mensaje, False, Date.Now, Archivo)
        taCorreos.Dispose()
    End Sub
    Public Function CTOD(ByVal cFecha As String) As Date
        Dim nDia, nMes, nYear As Integer
        nDia = Val(Strings.Right(cFecha, 2))
        nMes = Val(Mid(cFecha, 5, 2))
        nYear = Val(Strings.Left(cFecha, 4))
        CTOD = DateSerial(nYear, nMes, nDia)
    End Function

    Public Function DTOC(ByVal dFecha As Date) As String
        Dim cDia, cMes, cYear, sFecha As String
        sFecha = dFecha.ToShortDateString
        cDia = Strings.Left(sFecha, 2)
        cMes = Mid(sFecha, 4, 2)
        cYear = Strings.Right(sFecha, 4)
        DTOC = cYear & cMes & cDia
    End Function

    Public Function Stuff(ByVal Cadena As String, ByVal Lado As String, ByVal Llenarcon As String, ByVal Longitud As Integer) As String
        ' Declaración de variables de datos
        Dim cCadenaAuxiliar As String
        Dim nVeces As Integer
        Dim i As Integer

        nVeces = Longitud - Val(Len(Cadena))
        cCadenaAuxiliar = ""
        For i = 1 To nVeces
            cCadenaAuxiliar = cCadenaAuxiliar & Llenarcon
        Next
        If Lado = "D" Then
            Stuff = Cadena & cCadenaAuxiliar
        Else
            Stuff = cCadenaAuxiliar & Cadena
        End If
    End Function

    Public Function Leap(ByVal nYear As Integer) As Integer
        If nYear Mod 400 = 0 Then
            Leap = 1
        ElseIf nYear Mod 100 = 0 Then
            Leap = 0
        ElseIf nYear Mod 4 = 0 Then
            Leap = 1
        End If
    End Function

    Public Function MandaCorreoFase(De As String, Fase As String, Asunto As String, Mensaje As String, Optional ByVal Archivo As String = "") As Boolean
        Asunto = Asunto.Trim
        MandaCorreoFase = False
        taFASES.Fill(tFASES, Fase)
        For Each rFASES In tFASES.Rows
            taCorreos.Insert(De, rFASES.Correo, Asunto, Mensaje, False, Date.Now, Archivo)
            MandaCorreoFase = True
        Next
        taFASES.Dispose()
        Return MandaCorreoFase
    End Function
End Module
