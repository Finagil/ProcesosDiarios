Imports System.IO
Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient
Module Procesos
    Dim Arg() As String
    Sub Main()
        Arg = Environment.GetCommandLineArgs()
        If Arg.Length > 1 Then
            Select Case UCase(Arg(1))
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
                Case "PASIVOS"
                    Dim ID As Integer
                    Dim SoloUno As Boolean
                    If Arg.Length >= 3 Then
                        ID = Arg(2)
                        SoloUno = True
                    End If
                    'GeneraInteresesDiarios("2018-01-31")
                    'GeneraInteresesDiarios("2018-02-28")
                    'GeneraInteresesDiarios("2018-03-31")
                    If Date.Now.Day <= 6 Then
                        GeneraInteresesDiarios(Date.Now.Date.AddDays(Date.Now.Day * -1), SoloUno, ID) ' se procesa 6 dias lo del mes anterior
                    End If
                    GeneraInteresesDiarios(Date.Now.Date, SoloUno, ID) '.AddDays(Date.Now.Date.Day * -1))
            End Select
        End If
        Console.WriteLine("Terminado")
    End Sub
End Module
