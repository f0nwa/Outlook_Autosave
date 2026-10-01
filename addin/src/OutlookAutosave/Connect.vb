Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports OutlookAutosave.Core
Imports OutlookAutosave.Interop

<Assembly: ComVisible(False)>

''' <summary>
''' Точка входа COM-надстройки Outlook. Регистрируется установщиком в HKCU
''' (ProgID OutlookAutosave.Connect) и загружается Outlook при старте.
''' </summary>
<ComVisible(True)>
<Guid("6F0E8E54-3B7C-4E0B-9C51-2B7A3C1D5E10")>
<ProgId("OutlookAutosave.Connect")>
<ClassInterface(ClassInterfaceType.AutoDispatch)>
Public Class Connect
    Implements IDTExtensibility2, IRibbonExtensibility

    Public Const ClassId As String = "6F0E8E54-3B7C-4E0B-9C51-2B7A3C1D5E10"
    Public Const ProgIdValue As String = "OutlookAutosave.Connect"

    ''' <summary>ext_ConnectMode.ext_cm_AfterStartup: надстройка подключена после запуска Outlook.</summary>
    Private Const ExtConnectModeAfterStartup As Integer = 0

    Public Sub New()
    End Sub

#Region "IDTExtensibility2"

    Public Sub OnConnection(Application As Object, ConnectMode As Integer, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        Try
            OutlookHost.OutlookApp = Application
            AutoSaveEngine.Initialize()

            ' Надстройку включили в окне "Надстройки COM" при уже запущенном Outlook:
            ' OnStartupComplete в этом случае не вызывается, поэтому запускаемся сразу.
            If ConnectMode = ExtConnectModeAfterStartup Then
                AutoSaveEngine.Startup()
            End If
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: OnConnection failed: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        Try
            AutoSaveEngine.Startup()
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: OnStartupComplete failed: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        Try
            AutoSaveEngine.Shutdown()
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: OnBeginShutdown failed: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnDisconnection(RemoveMode As Integer, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        Try
            AutoSaveEngine.Shutdown()
            OutlookHost.OutlookApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: OnDisconnection failed: " & ex.ToString())
        End Try
    End Sub

#End Region

#Region "Лента"

    Public Function GetCustomUI(RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        If String.Equals(RibbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase) Then
            Return RibbonXml
        End If

        Return String.Empty
    End Function

    ' Обработчики вызываются лентой Office через IDispatch по имени метода.
    Public Sub OnSettingsClick(control As Object)
        Try
            UI.Commands.ShowSettings()
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: settings click failed: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnSaveSelectedClick(control As Object)
        Try
            UI.Commands.SaveSelected(OutlookHost.GetOwnerWindow())
        Catch ex As Exception
            Trace.WriteLine("OutlookAutosave: save selected click failed: " & ex.ToString())
        End Try
    End Sub

    Private Const RibbonXml As String =
        "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" &
        "<ribbon><tabs><tab idMso='TabMail'>" &
        "<group id='grpOutlookAutosave' label='Outlook Autosave'>" &
        "<button id='btnOasSettings' label='Настройки' size='large' imageMso='PropertySheet' onAction='OnSettingsClick' " &
        "screentip='Outlook Autosave' supertip='Статус, правила и настройки автосохранения вложений.'/>" &
        "<button id='btnOasSaveSelected' label='Сохранить выбранные' size='large' imageMso='FileSaveAs' onAction='OnSaveSelectedClick' " &
        "screentip='Сохранить выбранные' supertip='Сохранить вложения из выделенных писем в папку из настроек выбранных писем.'/>" &
        "</group>" &
        "</tab></tabs></ribbon>" &
        "<contextMenus>" &
        "<contextMenu idMso='ContextMenuMailItem'>" &
        "<button id='ctxOasSaveSelectedSingle' label='Сохранить вложения (Outlook Autosave)' imageMso='FileSaveAs' onAction='OnSaveSelectedClick'/>" &
        "</contextMenu>" &
        "<contextMenu idMso='ContextMenuMultipleItems'>" &
        "<button id='ctxOasSaveSelectedMultiple' label='Сохранить вложения (Outlook Autosave)' imageMso='FileSaveAs' onAction='OnSaveSelectedClick'/>" &
        "</contextMenu>" &
        "</contextMenus>" &
        "</customUI>"

#End Region

End Class
