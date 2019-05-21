Imports System.Runtime.InteropServices

Public Class KeyboardHook
    Implements IDisposable
    Private _HHookID As IntPtr
    'Private hInstance As IntPtr
    Private Hookstruct As KBDLLHOOKSTRUCT
    Private _dictionary As Dictionary(Of String, Integer)
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
    Public Sub New(ByVal dictionary As Dictionary(Of String, Integer))
        _HHookID = IntPtr.Zero
        _dictionary = dictionary
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
    Public Event CombKey(ByVal action As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal focus As String)
    Public Event PasteEvent(ByVal action As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal focus As String)
    'Public Event CombKey(ByVal action As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal focus As String)
    'Esta función se encarga de mostrar los eventos elegidos'
    Private Function KeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
        Dim action As Integer = 0
        Dim keyCode As Integer
        Dim focus As String = ""
        If nCode = HookCodes.HC_ACTION Then
            Select Case wParam
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    keyCode = CType(Marshal.PtrToStructure(lParam, Hookstruct.GetType()), KBDLLHOOKSTRUCT).vkCode
                    GrapCombKey(action, keyCode, focus)
            End Select
        End If
        Return CallNextHookEx(_HHookID, nCode, wParam, lParam)
    End Function
    'En este método se capturan las combinaciones de teclas más usadas'
    Private Sub GrapCombKey(ByRef action As Integer, keyCode As Integer, ByRef focus As String)
        'En esta parte se capturan las combinaciones que no se quiere registrar como evento de escritura o de comb de teclas'
        If GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_C Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_C _
            Or GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_X Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_X _
            Or GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_Z Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_Z _
            Or GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_Y Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_Y _
            Or GetAsyncKeyState(VK_LMENU) And keyCode = VK_TAB Or GetAsyncKeyState(VK_LMENU) And keyCode = VK_F4 _
            Or keyCode = VK_LCONTROL Or keyCode = VK_RCONTROL Or keyCode = VK_LMENU Then
            'do nothing'
        ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_V Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_V Or
            GetAsyncKeyState(VK_LSHIFT) And keyCode = VK_INSERT Or GetAsyncKeyState(VK_RSHIFT) And keyCode = VK_INSERT Then
            If Clipboard.ContainsText Or Clipboard.ContainsImage Then
                action = SearchValue(_dictionary, "Paste")
                focus = GetPathName()
                RaiseEvent PasteEvent(action, VK_LCONTROL, keyCode, focus)
            End If
        ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_S Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_S Then
                action = SearchValue(_dictionary, "CombCtrlS")
                focus = GetPathName()
                RaiseEvent CombKey(action, VK_LCONTROL, keyCode, focus)
            ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_G Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_G Then
                action = SearchValue(_dictionary, "CombCtrlG")
                focus = GetPathName()
                RaiseEvent CombKey(action, VK_LCONTROL, keyCode, focus)
            ElseIf GetAsyncKeyState(VK_LCONTROL) And keyCode = VK_F Or GetAsyncKeyState(VK_RCONTROL) And keyCode = VK_F Then
                action = SearchValue(_dictionary, "CombCtrlF")
                focus = GetPathName()
                RaiseEvent CombKey(action, VK_LCONTROL, keyCode, focus)
            Else
                focus = GetPathName()
            action = SearchValue(_dictionary, "Type")
            RaiseEvent KeyDown(action, focus)
        End If
    End Sub
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