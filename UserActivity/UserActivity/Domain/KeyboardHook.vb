Imports System.Runtime.InteropServices

Public Class KeyboardHook
    Implements IDisposable
    Private _HHookID As IntPtr
    'Private hInstance As IntPtr
    Private Hookstruct As KBDLLHOOKSTRUCT
    'Private _threadFocus As Threading.Thread
    '*******************************************************************************************'
    'Librerías DLL importadas'
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal HookProc As KBDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    'Se delega la función KBDLLHookProc de manera asíncrona y se declara el objeto de este tipo'
    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    'Getters'
    Public Property HHookID As IntPtr
        Get
            Return _HHookID
        End Get
        Set(value As IntPtr)
            _HHookID = value
        End Set
    End Property
    '********************************************************************************************'
    'Constructor'
    Public Sub New()
        _HHookID = IntPtr.Zero
        'hInstance = System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32
        _HHookID = SetWindowsHookEx(HookType.WH_KEYBOARD_LL, KBDLLHookProcDelegate, 0, 0)
    End Sub
    'Estructura KBLLHOOKSTRUCT obtiene información del ratón a bajo nivel'
    <StructLayout(LayoutKind.Sequential)>
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure
    <Flags()>
    Public Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum
    'Se declaran los eventos de teclado'
    Public Event KeyDown(ByVal action As Integer, ByVal focus As String)
    'Public Event CombKey(ByVal action As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal focus As String)
    'Esta función se encarga de mostrar los eventos elegidos'
    Private Function KeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
        Dim action As Integer
        Dim keyCode As Integer
        Dim focus As String
        If nCode = HookCodes.HC_ACTION Then
            Select Case wParam
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    'Capturamos las combinaciones de tecla más usadas'
                    keyCode = CType(Marshal.PtrToStructure(lParam, Hookstruct.GetType()), KBDLLHOOKSTRUCT).vkCode
                    If keyCode = VK_LCONTROL Or keyCode = VK_RCONTROL Or keyCode = VK_LMENU Or keyCode = VK_TAB Then
                        'do nothing'
                        'de esta manera no se registra una tecla perteneciente a una combinación'
                    Else
                        focus = GetPathName()
                        action = TypeAction.TypeApp
                        RaiseEvent KeyDown(action, focus)
                    End If
            End Select
        End If
        Return CallNextHookEx(_HHookID, nCode, wParam, lParam)
    End Function

    'Se utiliza esta interfaz para liberar la memoria asignada a un objeto administrado cuando ya no se utiliza ese objeto
    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not _HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
        MyBase.Finalize()
    End Sub
End Class