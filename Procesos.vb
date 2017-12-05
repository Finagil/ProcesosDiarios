﻿Imports System.IO
Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient
Module Procesos
    Dim Arg() As String
    Sub Main()
        Arg = Environment.GetCommandLineArgs()
        If Arg.Length > 1 Then
            Select Case UCase(Arg(1))
                Case "FACTURAS_CFDI"
                    Dim Fecha As Date = Today
                    'CFDI33.FacturarCFDI(Fecha, Arg(2))
                    CFDI33.FacturarCFDI_AV(Fecha)
                    CFDI33.FacturarCFDI(Fecha, "ANTERIORES")
                    CFDI33.FacturarCFDI(Fecha, "PREPAGO")
                    CFDI33.FacturarCFDI(Fecha, "DIA")

                Case "SALDOAVIO"
                    Console.WriteLine("Saldos Avios")
                    SaldosAvios()
                Case "BACKUPDB"
                    RepaldoDB()
                Case "DOMICILIACION"
                    EnviaLayout("B") 'Bancomer
                    EnviaLayout("O") 'otros bancos
                Case "GENERAPERSONAS"
                    PersonasHistoria.Main()
                Case "TERMINACONTRATO"
                    Termina_Contratos()
                Case "FACTORAJE"
                    Factoraje.NotificacionFactorajeFACT_VENC()
            End Select
        End If
        Console.WriteLine("Terminado")
    End Sub
End Module
