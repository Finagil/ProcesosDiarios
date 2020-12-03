Imports System.Data
Imports System.Math
Imports Microsoft.VisualBasic
Imports System.Net.Mail

Public Class Utilerias
    Public Sub CalcInte(ByVal drFacturas As DataRowCollection, ByVal drTasas As DataRowCollection, ByRef nTasaFact As Decimal, ByRef nDiasFact As Integer, ByRef nIntReal As Decimal, ByVal cFeven As String, ByVal cAnexo As String, ByVal cFechacon As String, ByVal cLetra As String, ByVal nSaldo As Decimal, ByVal cTipta As String, ByVal nDifer As Decimal)

        ' Declaración de variables de conexión

        Dim drFactura As DataRow

        ' Declaración de variables de datos

        Dim cAnterior As String
        Dim dAnterior As Date
        Dim dFeven As Date
        Dim nLetra As Byte

        nLetra = Val(cLetra)

        If nLetra = 1 Then
            cAnterior = cFechacon
            dAnterior = MGlobal.CTOD(cAnterior)
            dFeven = MGlobal.CTOD(cFeven)
            nDiasFact = DateDiff(DateInterval.Day, dAnterior, dFeven)
        Else
            For Each drFactura In drFacturas
                If cAnexo = drFactura("Anexo") And Val(drFactura("Letra")) = nLetra - 1 Then
                    cAnterior = drFactura("Feven")
                    dFeven = MGlobal.CTOD(cFeven)
                    dAnterior = MGlobal.CTOD(cAnterior)
                    nDiasFact = IIf(dAnterior < dFeven, DateDiff(DateInterval.Day, dAnterior, dFeven), 0)
                End If
            Next
        End If

        ' Cuando se trata del primer vencimiento de un contrato, se aplica la tasa y el diferencial pactados
        ' en las condiciones del contrato.   Por esta razón, solamente calcula la tasa de facturación para
        ' vencimientos posteriores al primero

        If nLetra > 1 Then
            If cTipta <> "7" Then
                nTasaFact = 0
                TraeTasa(drTasas, cTipta, cAnterior, nTasaFact, cFechacon)
            End If
        End If
        nTasaFact = nTasaFact + nDifer
        If nDiasFact < 0 Then
            nDiasFact = 0
            nIntReal = 0
        Else
            nIntReal = Round(nSaldo * nTasaFact / 36000 * nDiasFact, 2)
        End If

    End Sub

    Public Sub TraeTasa(ByVal drTasas As DataRowCollection, ByVal cTipta As String, ByVal cFeven As String, ByRef nTasaFact As Decimal, ByVal cFechacon As String)

        ' El parámetro drTasas contiene los valores de las diferentes tasas de interés.   En primera instancia
        ' contiene TODAS las tasas de interés del archivo HISTA.   Sin embargo, tiene asignada una Llave Primaria
        ' a fin de que la búsqueda se realice en forma directa en vez de secuencial.

        ' Declaración de variables de conexión ADO .NET

        Dim drTasa As DataRow

        For Each drTasa In drTasas

            If drTasa("Vigencia") = cFeven Then

                If cTipta = "1" And InStr(1, "34", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa Líder entre TIIP, TIIE

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "2" And InStr(1, "134", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa Líder entre CPP, TIIP, TIIE

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "3" And InStr(1, "123", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa Líder entre CPP, CETES, TIIP

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "4" And InStr(1, "1234", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa Líder entre CPP, CETES, TIIP, TIIE

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "5" And InStr(1, "5", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa NAFIN

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "6" And InStr(1, "4", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa TIIE

                    nTasaFact = IIf(drTasa("Valor") > nTasaFact, drTasa("Valor"), nTasaFact)

                ElseIf cTipta = "8" And InStr(1, "4", drTasa("Tasa"), CompareMethod.Text) > 0 Then

                    ' Tasa PROTEGIDA

                    Select Case cFechacon
                        Case Is >= "20020601"
                            nTasaFact = IIf(drTasa("Valor") < 13, drTasa("Valor"), 13)
                        Case Is >= "20020501"
                            nTasaFact = IIf(drTasa("Valor") < 12, drTasa("Valor"), 12)
                        Case Is >= "20020301"
                            nTasaFact = IIf(drTasa("Valor") < 13, drTasa("Valor"), 13)
                        Case Else
                            nTasaFact = IIf(drTasa("Valor") < 14, drTasa("Valor"), 14)
                    End Select

                End If

            End If

        Next

    End Sub

    Public Function CalcMora(ByVal cTipar As String, ByVal cTipo As String, ByVal cFecha As String, ByVal drUdis As DataRowCollection, ByVal nSaldo As Decimal, ByVal nTasaMoratoria As Decimal, ByVal nDiasMoratorios As Decimal, ByRef nMoratorios As Decimal, ByRef nIvaMoratorios As Decimal, ByVal nTasaIVACliente As Decimal) As Decimal

        ' Declaración de variables de datos

        Dim cFechaInicial As String
        Dim dFechaInicial As Date
        Dim nUdiFinal As Decimal
        Dim nUdiInicial As Decimal

        dFechaInicial = DateAdd(DateInterval.Day, -nDiasMoratorios, MGlobal.CTOD(cFecha))
        cFechaInicial = DTOC(dFechaInicial)
        nUdiInicial = 0
        nUdiFinal = 0

        nMoratorios = Round(nSaldo * nTasaMoratoria * (nDiasMoratorios) / 36000, 2)
        nIvaMoratorios = 0

        ' Hasta el 10 de enero de 2010 se calculaba el IVA de los moratorios en base a UDIS sin importar el tipo de financiamiento lo cual era incorrecto.
        ' A partir del 11 de enero solo existe IVA moratorios para :
        ' Arrendamiento Financiero (en base a UDIS) y para
        ' Crédito Refaccionario o Crédito Simple siempre y cuando se trate de una Persona Física SIN actividad empresarial en cuyo caso será igual al Porcentaje de IVA vigente 

        If cTipar = "F" Then
            nIvaMoratorios = CalcIvaU(drUdis, nSaldo, nTasaMoratoria, cFechaInicial, cFecha, nUdiInicial, nUdiFinal, (nTasaIVACliente / 100))
        Else
            If cTipo = "F" Then
                'If IVA_Interes_TasaReal = False Or cFecha < "20160101" Then 'Enterar IVA Basado en fujo = TRUE o direco sobre base nominal = False #ECT20151015.n
                '    nIvaMoratorios = Round(nMoratorios * (nTasaIVACliente / 100), 2)
                'Else
                '    nIvaMoratorios = CalcIvaU(drUdis, nSaldo, nTasaMoratoria, cFechaInicial, cFecha, nUdiInicial, nUdiFinal, (nTasaIVACliente / 100))
                'End If
            End If
        End If

    End Function

    Public Function CalcIvaU(ByVal drUdis As DataRowCollection, ByVal nSaldo As Decimal, ByVal nTasa As Decimal, ByVal cFechaInicial As String, ByVal cFechaFinal As String, ByRef nUdiInicial As Decimal, ByRef nUdiFinal As Decimal, ByVal nPorcentajeIVA As Decimal) As Decimal

        ' Declaración de variables de datos

        Dim drUdi As DataRow
        Dim dFechaInicial As Date
        Dim dFechaFinal As Date
        Dim nDias As Integer

        nUdiInicial = 0
        nUdiFinal = 0
        CalcIvaU = 0

        If nSaldo > 0 Then

            dFechaInicial = MGlobal.CTOD(cFechaInicial)
            dFechaFinal = MGlobal.CTOD(cFechaFinal)
            nDias = DateDiff(DateInterval.Day, dFechaInicial, dFechaFinal)

            If nDias > 0 Then
                dFechaInicial = DateAdd(DateInterval.Day, -1, dFechaInicial)
                dFechaFinal = DateAdd(DateInterval.Day, -1, dFechaFinal)
                cFechaInicial = DTOC(dFechaInicial)
                cFechaFinal = DTOC(dFechaFinal)
                For Each drUdi In drUdis
                    If drUdi("Vigencia") = cFechaInicial Then
                        nUdiInicial = drUdi("Udi")
                    End If
                    If drUdi("Vigencia") = cFechaFinal Then
                        nUdiFinal = drUdi("Udi")
                    End If
                Next
                If nUdiFinal <= nUdiInicial Then
                    CalcIvaU = nSaldo * nTasa * nDias / 36000 * nPorcentajeIVA
                Else
                    CalcIvaU = nSaldo * ((nTasa * nDias / 36000) - ((nUdiFinal / nUdiInicial) - 1)) * nPorcentajeIVA
                End If
                CalcIvaU = Round(CalcIvaU, 2)
                If CalcIvaU < 0 Then
                    CalcIvaU = 0
                End If
            End If
        End If

    End Function

    'Public Function DiaSemana(ByVal dFecha As Date)

    '    Dim nDay As Byte
    '    Dim nYear As Integer
    '    Dim nMonth As Byte
    '    Dim nAños As Integer
    '    Dim nAñosb As Integer
    '    Dim nLeap As Byte
    '    Dim i As Integer
    '    Dim nMes As Integer
    '    Dim nDia As Integer

    '    nDay = Day(dFecha)
    '    nMonth = Month(dFecha)
    '    nYear = Year(dFecha)

    '    nAños = nYear - 1933
    '    nLeap = 0
    '    nAñosb = 0

    '    For i = 1933 To nYear
    '        nLeap = Leap(i)
    '        If nLeap = 1 Then
    '            nAñosb += 1
    '            nLeap = 0
    '        End If
    '    Next

    '    Select Case nMonth
    '        Case 1, 10
    '            nMes = 0
    '        Case 2, 3, 11
    '            nMes = 3
    '        Case 4, 7
    '            nMes = 6
    '        Case 5
    '            nMes = 1
    '        Case 6
    '            nMes = 4
    '        Case 8
    '            nMes = 2
    '        Case 9, 12
    '            nMes = 5
    '    End Select

    '    nDia = (nAños + nAñosb + nMes + nDay) Mod 7
    '    If nDia = 1 Then
    '        dFecha = DateAdd(DateInterval.Day, 1, dFecha)
    '    ElseIf nDia = 0 Then
    '        dFecha = DateAdd(DateInterval.Day, 2, dFecha)
    '    End If

    '    DiaSemana = dFecha.ToShortDateString

    'End Function



End Class
