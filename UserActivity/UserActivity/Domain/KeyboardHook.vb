Imports System.Runtime.InteropServices
Imports System.Text

Public Class KeyboardHook
    Implements IDisposable
    Private _HHookID As IntPtr
    Private hInstance As IntPtr
    Private Hookstruct As KBDLLHOOKSTRUCT
    Private _focus As String
    '*******************************************************************************************'
    'Librerías DLL importadas'
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal HookProc As KBDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    '********************************************************************************************'
    'Se delega la función KBDLLHookProc y se declara el objeto de este tipo'
    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    'Getter'
    Public Property HHookID As IntPtr
        Get
            Return _HHookID
        End Get
        Set(value As IntPtr)
            _HHookID = value
        End Set
    End Property
    'Constructor'
    Public Sub New(ByRef focusMem As String)
        _HHookID = IntPtr.Zero
        _focus = focusMem
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
    Public Event KeyDown(ByVal focus As String, ByVal procID As Integer)
    Public Event CombKey(ByVal key As Keys, ByVal vKey As Integer, ByVal focus As String, ByVal procID As Integer)
    'Método que devuelve el foco y por lo tanto el titulo de la aplicación'
    Private Function GetForegroundInfo()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim length As Integer
        Dim wTitle As StringBuilder

        If hWnd <> IntPtr.Zero Then
            length = GetWindowTextLength(hWnd)
            wTitle = New System.Text.StringBuilder("", length + 1)
            If length > 0 Then
                GetWindowText(hWnd, wTitle, wTitle.Capacity)
            End If
            Return wTitle.ToString
        End If
    End Function
    'Numero de Proceso'
    Private Function GetProcessID()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        Dim wProcID As Integer = Nothing
        procID = GetWindowThreadProcessId(hWnd, wProcID)
        Return procID
    End Function
    'Esta función se encarga de mostrar los eventos elegidos'
    Private Function KeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
        Dim processID As Integer
        Dim keyCode As Integer
        Dim flag As Integer
        Dim focus As String
        If nCode = HookCodes.HC_ACTION Then
            keyCode = CType(Marshal.PtrToStructure(lParam, Hookstruct.GetType()), KBDLLHOOKSTRUCT).vkCode
            flag = CType(Marshal.PtrToStructure(lParam, Hookstruct.GetType()), KBDLLHOOKSTRUCT).flags
            'focus = GetForegroundInfo()
            Select Case wParam
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    If GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_C Then
                        RaiseEvent CombKey(VK_LCONTROL, keyCode, focus, processID)
                    ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_V Then
                        RaiseEvent CombKey(VK_LCONTROL, keyCode, focus, processID)
                    ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_X Then
                        RaiseEvent CombKey(VK_LCONTROL, keyCode, focus, processID)
                    ElseIf GetAsyncKeyState(VK_LMENU) And keyCode = VK_TAB Then
                        focus = GetPathName()
                        'If focus <> _focus Then
                        processID = GetProcessID()
                        RaiseEvent CombKey(VK_LMENU, keyCode, focus, processID)
                        '    _focus = focus
                        'Else
                        'do nothing'
                        'End If
                    ElseIf GetAsyncKeyState(VK_LMENU) And keyCode = VK_F4 Then
                        RaiseEvent CombKey(VK_LMENU, keyCode, focus, processID)
                    ElseIf keyCode = VK_LCONTROL Or keyCode = VK_LMENU Or keyCode = VK_TAB Then
                        'do nothing'
                        'de esta manera no se registra una tecla perteneciente a una combinación'
                    Else
                        processID = GetProcessID()
                        RaiseEvent KeyDown(focus, processID)
                    End If
            End Select
        End If
        Return CallNextHookEx(_HHookID, nCode, wParam, lParam)
    End Function
    'Path Completo'
    Private Function GetPathName()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim proc As Process
        Dim wProcID As Integer = Nothing
        Dim wFileName As String = ""

        If hWnd <> IntPtr.Zero Then
            GetWindowThreadProcessId(hWnd, wProcID)
            proc = Process.GetProcessById(wProcID)
            'por si alguno no tiene permisos de lectura'
            Try
                wFileName = proc.MainModule.FileName
            Catch ex As Exception
                wFileName = ""
            End Try
        End If
        Return wFileName
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