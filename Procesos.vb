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
                    Dim Dias As Integer
                    If Arg.Length >= 3 Then
                        Dias = Arg(2)
                    Else
                        Dias = 0
                    End If
                    EnviaLayout("B", Dias) 'Bancomer
                    EnviaLayout("O", Dias) 'otros bancos
                Case "GENERAPERSONAS"
                    PersonasHistoria.Main()
                Case "TERMINACONTRATO"
                    Termina_Contratos()
                Case "FACTORAJE"
                    Factoraje.NotificacionFactorajeFACT_VENC()
                Case "PASIVOS"
                    Dim ID As Integer
                    If Arg.Length >= 3 Then
                        ID = Arg(2)
                    End If
                    'GeneraInteresesDiarios("2017-12-31", ID)
                    'GeneraInteresesDiarios("2018-01-31", ID)
                    'GeneraInteresesDiarios("2018-02-28", ID)
                    'GeneraInteresesDiarios("2018-03-31", ID)
                    'GeneraInteresesDiarios("2018-04-30", ID)
                    'GeneraInteresesDiarios("2018-05-31", ID)
                    'GeneraInteresesDiarios("2018-06-30", ID)
                    'GeneraInteresesDiarios("2018-07-31", ID)
                    'GeneraInteresesDiarios("2018-08-31", ID)
                    'GeneraInteresesDiarios("2019-01-31", ID)
                    If Date.Now.Day <= 6 Then
                        GeneraInteresesDiarios(Date.Now.Date.AddDays(Date.Now.Day * -1), ID) ' se procesa 6 dias lo del mes anterior
                    End If
                    GeneraInteresesDiarios(Date.Now.Date, ID) '.AddDays(Date.Now.Date.Day * -1))
            End Select
        End If
        Console.WriteLine("Terminado")
    End Sub
End Module
