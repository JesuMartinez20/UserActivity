Imports System.Runtime.InteropServices
Imports System.Text

Public Class MouseHook
    Implements IDisposable
    Private hInstance As Integer
    Private _HHookID As IntPtr
    Private Hookstruct As MSLLHOOKSTRUCT
    Private _focus As String
    '************************************************************************************************'
    'Librerías DLL importadas'
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function SetWindowsHookEx(ByVal idHook As HookType, ByVal HookProc As MSDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    '***********************************************************************************************'
    'Se delega la función KBDLLHookProc y se declara el objeto de este tipo'
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
    'Constructor'
    Public Sub New(ByRef focus As String)
        _focus = focus
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

    Public Enum WheelDirection
        WheelUp
        WheelDown
    End Enum
    'Eventos de ratón'
    Public Event MouseLeftDown(ByVal focus As String, ByVal procID As Integer)
    Public Event MouseRightDown(ByVal location As Point)
    Public Event MouseWheel(ByVal focus As String, ByVal procID As Integer)
    Public Event MouseDoubleClick(ByVal focus As String)
    'Método que devuelve el foco y por lo tanto el titulo de la aplicación'
    Private Function GetForegroundInfo()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim length As Integer
        Dim wTitle As StringBuilder = New System.Text.StringBuilder("", 0)

        If hWnd <> IntPtr.Zero Then
            length = GetWindowTextLength(hWnd)
            wTitle = New System.Text.StringBuilder("", length + 1)
            If length > 0 Then
                GetWindowText(hWnd, wTitle, wTitle.Capacity)
            End If
        End If
        Return wTitle.ToString
    End Function
    'Numero de Proceso'
    Private Function GetProcessID()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        Dim wProcID As Integer = Nothing
        procID = GetWindowThreadProcessId(hWnd, wProcID)
        Return procID
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
                wFileName = Proc.MainModule.FileName
            Catch ex As Exception
                wFileName = ""
            End Try
        End If
        Return wFileName
    End Function
    'Esta función se encarga de mostrar los eventos elegidos'
    Private Function MouseProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
        Dim processID As Integer
        'Dim mouseData As Integer
        Dim focus As String = ""
        'Dim wDirection As WheelDirection
        If nCode = HookCodes.HC_ACTION Then
            Select Case wParam
                Case WM_LBUTTONDOWN
                    focus = GetPathName()
                    'If focus <> _focus Then
                    processID = GetProcessID()
                    RaiseEvent MouseLeftDown(focus, processID)
                    '_focus = focus
            'Else
            'do nothing'
        'End If
                Case WM_RBUTTONDOWN
                    'Para el clipboard'
                Case WM_MOUSEWHEEL
                    'mouseData = CType(Marshal.PtrToStructure(lParam, Hookstruct.GetType()), MSLLHOOKSTRUCT).mouseData
                    'wDirection = DirectionWheel(mouseData, wDirection)
                    Static focusMemWheel As String = GetPathName()
                    Static focusInitializeWheel = GetPathName()
                    focus = GetPathName()
                    If focus = focusInitializeWheel Then
                        processID = GetProcessID()
                        RaiseEvent MouseWheel(focus, processID)
                        focusInitializeWheel = Nothing
                    ElseIf (focus <> focusMemWheel) Then
                        processID = GetProcessID()
                        RaiseEvent MouseWheel(focus, processID)
                        focusMemWheel = focus
                    Else
                        Debug.WriteLine("nothing")
                    End If
            End Select
        End If
        Return CallNextHookEx(_HHookID, nCode, wParam, lParam)
    End Function

    Private Shared Function DirectionWheel(mouseData As Integer, wDirection As WheelDirection) As WheelDirection
        If mouseData < 0 Then
            wDirection = WheelDirection.WheelDown
        Else
            wDirection = WheelDirection.WheelUp
        End If
        Return wDirection
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
