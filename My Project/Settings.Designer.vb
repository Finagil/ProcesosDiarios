﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.9.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "Funcionalidad para autoguardar My.Settings"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=server-raid2;Initial Catalog=production;Persist Security Info=True;Us"& _ 
            "er ID=User_PRO;Password=User_PRO2015")>  _
        Public ReadOnly Property ConnectionString() As String
            Get
                Return CType(Me("ConnectionString"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-raid2\Puente Sistemas\")>  _
        Public ReadOnly Property RutaOrigenBackup_Raid() As String
            Get
                Return CType(Me("RutaOrigenBackup_Raid"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-nas\RespaldosDB\")>  _
        Public Property RutaBackupDB() As String
            Get
                Return CType(Me("RutaBackupDB"),String)
            End Get
            Set
                Me("RutaBackupDB") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-nas\RespaldoServerRaid\")>  _
        Public Property RutaBackupRaid() As String
            Get
                Return CType(Me("RutaBackupRaid"),String)
            End Get
            Set
                Me("RutaBackupRaid") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-nas\RespaldoServerMinds\")>  _
        Public Property RutaBackupMinds() As String
            Get
                Return CType(Me("RutaBackupMinds"),String)
            End Get
            Set
                Me("RutaBackupMinds") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-nas\RespaldoServerContpaq\DataBases")>  _
        Public Property RutaBackupContpaq() As String
            Get
                Return CType(Me("RutaBackupContpaq"),String)
            End Get
            Set
                Me("RutaBackupContpaq") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\COMPAC01\Diario\")>  _
        Public Property RutaOrigenBackup_Contpaq() As String
            Get
                Return CType(Me("RutaOrigenBackup_Contpaq"),String)
            End Get
            Set
                Me("RutaOrigenBackup_Contpaq") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=SERVER-RAID2;Initial Catalog=production;User ID=User_PRO;Password=Use"& _ 
            "r_PRO2015")>  _
        Public Property ConnectionStringDOMI() As String
            Get
                Return CType(Me("ConnectionStringDOMI"),String)
            End Get
            Set
                Me("ConnectionStringDOMI") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=server-raid2;Initial Catalog=WEB_Finagil;Persist Security Info=True;U"& _ 
            "ser ID=User_PRO;Password=User_PRO2015")>  _
        Public ReadOnly Property ConnectionStringWEB() As String
            Get
                Return CType(Me("ConnectionStringWEB"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("X:\EDGAR\Fuentes\")>  _
        Public Property RutaOrigenFuentes() As String
            Get
                Return CType(Me("RutaOrigenFuentes"),String)
            End Get
            Set
                Me("RutaOrigenFuentes") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=server-raid2;Initial Catalog=Production;Persist Security Info=True;Us"& _ 
            "er ID=User_PRO;Password=User_PRO2015")>  _
        Public ReadOnly Property ConectionStringCFDI() As String
            Get
                Return CType(Me("ConectionStringCFDI"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\server-nas\FacturasCFDI\FoliosEkomercio\")>  _
        Public Property RutaFolios() As String
            Get
                Return CType(Me("RutaFolios"),String)
            End Get
            Set
                Me("RutaFolios") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.ProcesosDiarios.My.MySettings
            Get
                Return Global.ProcesosDiarios.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
