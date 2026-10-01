Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

' Объявления COM-интерфейсов, которые обычно берутся из extensibility.dll и office.dll.
' Они объявлены здесь, чтобы надстройка не зависела от Office PIA и собиралась без Visual Studio.
Namespace Interop

    ''' <summary>Интерфейс COM-надстройки Office (extensibility.dll).</summary>
    <ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),
     TypeLibType(TypeLibTypeFlags.FDual Or TypeLibTypeFlags.FDispatchable)>
    Public Interface IDTExtensibility2
        <DispId(1)>
        Sub OnConnection(<[In], MarshalAs(UnmanagedType.IDispatch)> Application As Object,
                         <[In]> ConnectMode As Integer,
                         <[In], MarshalAs(UnmanagedType.IDispatch)> AddInInst As Object,
                         <[In], MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <DispId(2)>
        Sub OnDisconnection(<[In]> RemoveMode As Integer,
                            <[In], MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <DispId(3)>
        Sub OnAddInsUpdate(<[In], MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <DispId(4)>
        Sub OnStartupComplete(<[In], MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <DispId(5)>
        Sub OnBeginShutdown(<[In], MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    End Interface

    ''' <summary>Интерфейс ленты Office (office.dll, Microsoft.Office.Core.IRibbonExtensibility).</summary>
    <ComImport, Guid("000C0396-0000-0000-C000-000000000046"),
     TypeLibType(TypeLibTypeFlags.FDual Or TypeLibTypeFlags.FDispatchable)>
    Public Interface IRibbonExtensibility
        <DispId(1)>
        Function GetCustomUI(<[In], MarshalAs(UnmanagedType.BStr)> RibbonID As String) As <MarshalAs(UnmanagedType.BStr)> String
    End Interface

    ''' <summary>Позволяет получить HWND окна Outlook, чтобы наши окна были модальными относительно него.</summary>
    <ComImport, Guid("00000114-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IOleWindow
        Sub GetWindow(<Out> ByRef phwnd As IntPtr)
        Sub ContextSensitiveHelp(<[In], MarshalAs(UnmanagedType.Bool)> fEnterMode As Boolean)
    End Interface

    ''' <summary>Элемент оболочки Windows (для современного диалога выбора папки).</summary>
    <ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellItem
        Sub BindToHandler(<[In]> pbc As IntPtr, <[In]> ByRef bhid As Guid, <[In]> ByRef riid As Guid, <Out> ByRef ppv As IntPtr)
        Sub GetParent(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
        Sub GetDisplayName(<[In]> sigdnName As UInteger, <Out, MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String)
        Sub GetAttributes(<[In]> sfgaoMask As UInteger, <Out> ByRef psfgaoAttribs As UInteger)
        Sub Compare(<[In], MarshalAs(UnmanagedType.Interface)> psi As IShellItem, <[In]> hint As UInteger, <Out> ByRef piOrder As Integer)
    End Interface

    ''' <summary>IFileDialog (включая метод Show из IModalWindow в начале таблицы).</summary>
    <ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileDialog
        <PreserveSig>
        Function Show(<[In]> hwndOwner As IntPtr) As Integer
        Sub SetFileTypes(<[In]> cFileTypes As UInteger, <[In]> rgFilterSpec As IntPtr)
        Sub SetFileTypeIndex(<[In]> iFileType As UInteger)
        Sub GetFileTypeIndex(<Out> ByRef piFileType As UInteger)
        Sub Advise(<[In]> pfde As IntPtr, <Out> ByRef pdwCookie As UInteger)
        Sub Unadvise(<[In]> dwCookie As UInteger)
        Sub SetOptions(<[In]> fos As UInteger)
        Sub GetOptions(<Out> ByRef pfos As UInteger)
        Sub SetDefaultFolder(<[In], MarshalAs(UnmanagedType.Interface)> psi As IShellItem)
        Sub SetFolder(<[In], MarshalAs(UnmanagedType.Interface)> psi As IShellItem)
        Sub GetFolder(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
        Sub GetCurrentSelection(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
        Sub SetFileName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String)
        Sub GetFileName(<Out, MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String)
        Sub SetTitle(<[In], MarshalAs(UnmanagedType.LPWStr)> pszTitle As String)
        Sub SetOkButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszText As String)
        Sub SetFileNameLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> pszLabel As String)
        Sub GetResult(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem)
        Sub AddPlace(<[In], MarshalAs(UnmanagedType.Interface)> psi As IShellItem, <[In]> fdap As Integer)
        Sub SetDefaultExtension(<[In], MarshalAs(UnmanagedType.LPWStr)> pszDefaultExtension As String)
        Sub Close(<[In], MarshalAs(UnmanagedType.Error)> hr As Integer)
        Sub SetClientGuid(<[In]> ByRef guid As Guid)
        Sub ClearClientData()
        Sub SetFilter(<[In]> pFilter As IntPtr)
    End Interface

    ''' <summary>
    ''' Исходящий интерфейс событий Outlook.Items (ItemsEvents, msoutl.olb). Объявлен здесь,
    ''' чтобы подписываться на ItemAdd без Outlook PIA.
    ''' </summary>
    <ComVisible(True), Guid("00063077-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
    Public Interface IItemsEvents
        <DispId(&HF001)>
        Sub ItemAdd(<[In], MarshalAs(UnmanagedType.IDispatch)> Item As Object)

        <DispId(&HF002)>
        Sub ItemChange(<[In], MarshalAs(UnmanagedType.IDispatch)> Item As Object)

        <DispId(&HF003)>
        Sub ItemRemove()
    End Interface

    Friend NotInheritable Class NativeMethods
        Private Sub New()
        End Sub

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
        Friend Shared Sub SHCreateItemFromParsingName(<[In], MarshalAs(UnmanagedType.LPWStr)> pszPath As String,
                                                      <[In]> pbc As IntPtr,
                                                      <[In]> ByRef riid As Guid,
                                                      <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IShellItem)
        End Sub
    End Class

End Namespace
