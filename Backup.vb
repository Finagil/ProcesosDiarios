Imports System.IO
Imports System.IO.Compression
Module Backup

    Public Sub RepaldoDB()
        MGlobal.EnviacORREO("ecacerest@finagil.com.mx", Date.Now, "Inicio BackupDB", "BackupDB@Finagil.com.mx")
        Dim folder As New DirectoryInfo(My.Settings.RutaOrigenBackup_Raid)
        Dim folder2 As New DirectoryInfo(My.Settings.RutaBackupDB)
        Dim Inicio As String = ""
        Dim Aux As Integer = 0
        Try
            '*****Respalda Fuentes
            'Console.WriteLine("Los subdirectorios de {0}", My.Settings.RutaOrigenFuentes)
            'recorrerDir(My.Settings.RutaOrigenFuentes, 0, My.Settings.RutaBackupDB, False)
            'recorrerDir(My.Settings.RutaOrigenFuentes, 0, My.Settings.RutaBackupRaid, True)
            '*****Respalda Fuentes

            For Each file1 As FileInfo In folder.GetFiles("*.bak")
                Aux = InStr(file1.Name, "_")
                Inicio = Mid(file1.Name, 1, Aux)
                For Each file2 As FileInfo In folder2.GetFiles(Inicio & "*.bak")
                    Console.WriteLine("Borrando: " & file2.FullName)
                    file2.Delete()
                Next
                Console.WriteLine("Copia 1 : " & file1.FullName)
                file1.CopyTo(My.Settings.RutaBackupDB & file1.Name, True)
                Console.WriteLine("Copia 2: " & file1.FullName)
                file1.CopyTo(My.Settings.RutaBackupRaid & file1.Name, True)
                Console.WriteLine("Borrando: " & file1.FullName)
                file1.Delete()
            Next

            folder = New DirectoryInfo(My.Settings.RutaOrigenBackup_Contpaq)
            For Each File1 As FileInfo In folder.GetFiles("*.bak")

                Aux = InStr(File1.Name, "_")
                Inicio = Mid(File1.Name, 1, Aux - 1)
                For Each file2 As FileInfo In folder2.GetFiles(Inicio & "*.bak")
                    Console.WriteLine("Borrando: " & file2.FullName)
                    file2.Delete()
                Next
                Console.WriteLine("Copia 1 : " & File1.FullName)
                File1.CopyTo(My.Settings.RutaBackupDB & File1.Name, True)
                Console.WriteLine("Copia 2: " & File1.FullName)
                File1.CopyTo(My.Settings.RutaBackupContpaq & File1.Name, True)
                Console.WriteLine("Borrando: " & File1.FullName)
                File1.Delete()
            Next

            folder = New DirectoryInfo(My.Settings.RutaBackupDB)
            For Each File1 As FileInfo In folder.GetFiles("minds*.bak")
                Console.WriteLine("Copia 2: " & File1.FullName)
                File1.CopyTo(My.Settings.RutaBackupMinds & File1.Name, True)
            Next
            For Each File1 As FileInfo In folder.GetFiles("Preven*.bak")
                Console.WriteLine("Copia 2: " & File1.FullName)
                File1.CopyTo(My.Settings.RutaBackupMinds & File1.Name, True)
            Next

        Catch ex As Exception
            MGlobal.EnviacORREO("ecacerest@finagil.com.mx", ex.Message, "Error BackupDB", "BackupDB@Finagil.com.mx")
        Finally
            MGlobal.EnviacORREO("ecacerest@finagil.com.mx", Date.Now, "Fin BackupDB", "BackupDB@Finagil.com.mx")
        End Try
    End Sub

    Public Sub Compress(directorySelected As DirectoryInfo)
        'Shell("cmd.exe /k NET USE \\192.168.10.231\Projects\Production /user:Agil\desarrollo 515t3m45x")

        For Each fileToCompress As FileInfo In directorySelected.GetFiles()
            Using originalFileStream As FileStream = fileToCompress.OpenRead()
                If (File.GetAttributes(fileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden And fileToCompress.Extension <> ".gz" Then
                    Using compressedFileStream As FileStream = File.Create(My.Settings.RutaOrigenBackup_Raid & "\" & fileToCompress.Name & ".gz")
                        Using compressionStream As New GZipStream(compressedFileStream, CompressionMode.Compress)


                            originalFileStream.CopyTo(compressionStream)
                        End Using
                    End Using
                    Dim info As New FileInfo(My.Settings.RutaOrigenBackup_Raid & "\" & fileToCompress.Name & ".gz")
                    Console.WriteLine("Compressed {0} from {1} to {2} bytes.", fileToCompress.Name,
                                      fileToCompress.Length.ToString(), info.Length.ToString())

                End If
            End Using
        Next
    End Sub



    Private Sub recorrerDir(ByVal elDir As String, ByVal nivel As Integer, ByVal RutaNueva As String, ByVal ConFec As Boolean)
        ' La sangría del nivel examinado
        Try


            Dim sangria As String = New String(" "c, nivel)
            Dim infoReader1 As System.IO.FileInfo
            Dim infoReader2 As System.IO.FileInfo
            Dim r As String = ""
            Dim rr As String = ""


            ' Los subdirectorios del directorio indicado
            Dim directorios As String()
            If nivel = 0 Then
                directorios = Directory.GetDirectories(elDir, "production")
            Else
                directorios = Directory.GetDirectories(elDir)
            End If
            Console.Write("{0}Directorio {1} con {2} subdirectorios", sangria, elDir, directorios.Length)
            Dim ficheros As String() = Directory.GetFiles(elDir)
            Console.WriteLine(" y {0} ficheros", ficheros.Length)

            ' Si tiene subdirectorios, recorrerlos
            If directorios.Length > 0 Then
                For Each d As String In directorios
                    ' Llamar de forma recursiva a este mismo método
                    'crea las carpetas
                    If ConFec Then
                        r = RutaNueva & Mid(d, elDir.Length + 1, d.Length) & Date.Now.ToString("ddMMyy") & "\"
                    Else
                        r = RutaNueva & Mid(d, elDir.Length + 1, d.Length) & "\"
                    End If

                    If InStr(r, "\bin\") > 0 Or InStr(r, "\obj\") > 0 Or InStr(r, "\.") > 0 Then
                        ' no pasamos estas carpetas
                    Else
                        If Not Directory.Exists(r) Then
                            Directory.CreateDirectory(r)
                        End If
                        recorrerDir(d, nivel + 2, r, False)
                    End If

                Next
            End If

            ' Si tiene archivos, recorrerlos
            If ficheros.Length > 0 Then
                For Each f As String In ficheros
                    'copia los archivos
                    'If InStr(UCase(f), ".JPG") > 0 Or InStr(UCase(f), ".TIF") > 0 Or InStr(UCase(f), ".PDF") > 0 Then
                    rr = RutaNueva & Mid(f, elDir.Length + 1, f.Length)
                    If Not File.Exists(rr) Then
                        File.Copy(f, rr, True)
                    Else
                        infoReader1 = My.Computer.FileSystem.GetFileInfo(f)
                        infoReader2 = My.Computer.FileSystem.GetFileInfo(rr)
                        If infoReader1.Length <> infoReader2.Length Then
                            File.Copy(f, rr, True)
                        Else
                            File.Copy(f, rr, True)
                        End If
                    End If
                    'End If
                    Console.WriteLine("       Ficheros: {0}", f)
                Next
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            MGlobal.EnviacORREO("Ecacerest@lamoderna.com.mx", ex.Message, "Error de Onbase " & Date.Now, "ProcesosDiarios@Finagil.com.mx")
        End Try
    End Sub


End Module
