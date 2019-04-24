Imports System.Runtime.InteropServices

Public Class MouseHook
    Implements IDisposable
    'Private hInstance As Integer
    Private _HHookID As IntPtr
    Private Hookstruct As MSLLHOOKSTRUCT
    '************************************************************************************************'
    'Librerías DLL importadas'
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal HookProc As MSDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    'Se delega la función KBDLLHookProc de manera asíncrona y se declara el objeto de este tipo'
    Private Delegate Function MSDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private MSDLLHookProcDelegate As MSDLLHookProc = New MSDLLHookProc(AddressOf MouseProc)
    'Getter'
    Public Property HHookID As IntPtr
        Get
            Return _HHookID
        End Get
        Set(value As IntPtr)
            _HHookID = value
        End Set
    End Property
    '***********************************************************************************************'
    'Constructor'
    Public Sub New()
        _HHookID = IntPtr.Zero
        'hInstance = System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32
        _HHookID = SetWindowsHookEx(HookType.WH_MOUSE_LL, MSDLLHookProcDelegate, 0, 0)
    End Sub
    'Estructura MSLLHOOKSTRUCT obtiene información del ratón a bajo nivel'
    <StructLayout(LayoutKind.Sequential)>
    Public Structure MSLLHOOKSTRUCT
        Public pt As Point
        Public mouseData As Int32
        Public flags As MSLLHOOKSTRUCTFlags
        Public time As Int32
        Public dwExtraInfo As UIntPtr
    End Structure
    <Flags()>
    Public Enum MSLLHOOKSTRUCTFlags As Int32
        LLMHF_INJECTED = 1
    End Enum
    'Eventos de ratón'
    Public Event MouseLeftDown(ByVal focus As String)
    Public Event MouseWheel(ByVal action As Integer, ByVal focus As String)
    'Esta función se encarga de levantar los eventos elegidos'
    Private Function MouseProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
        Dim focus As String = ""
        Dim action As Integer
        If nCode = HookCodes.HC_ACTION Then
            Select Case wParam
                Case WM_MOUSEWHEEL
                    focus = GetPathName()
                    action = TypeAction.ScrollApp
                    RaiseEvent MouseWheel(action, focus)
            End Select
        End If
        Return CallNextHookEx(_HHookID, nCode, wParam, lParam)
    End Function
#Region "IDisposable Support"
    ' TODO: reemplace Finalize() solo si el anterior Dispose(disposing As Boolean) tiene código para liberar recursos no administrados.
    Protected Overrides Sub Finalize()
        If Not _HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(_HHookID)
        End If
        MyBase.Finalize()
    End Sub
    ' Visual Basic agrega este código para implementar correctamente el patrón descartable.
    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(_HHookID)
        End If
    End Sub
#End Region
End Class
