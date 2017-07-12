Module mdlImpresoras

    Public Const HWND_BROADCAST = &HFFFF
    Public Const WM_WININICHANGE = &H1A

    ' constants for DEVMODE structure
    Public Const CCHDEVICENAME = 32
    Public Const CCHFORMNAME = 32

    ' constants for DesiredAccess member of PRINTER_DEFAULTS
    Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Public Const PRINTER_ACCESS_ADMINISTER = &H4
    Public Const PRINTER_ACCESS_USE = &H8
    Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
    PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

    ' constant that goes into PRINTER_INFO_5 Attributes member
    ' to set it as default
    Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

    ' Constant for OSVERSIONINFO.dwPlatformId
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2

    Public Const GW_CHILD = 5

    Public Declare Function GetClassName Lib "user32" Alias _
        "GetClassNameA" (ByVal hwnd As Int32, ByVal lpString As String, _
        ByVal cch As Int32) As Int32

    Public Const GW_HWNDNEXT = 2

    Public Structure RECT
        Dim Left As Int32
        Dim Top As Int32
        Dim Right As Int32
        Dim Bottom As Int32
    End Structure

    Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Int32, ByRef lpRect As RECT) As Int32

    Public Structure OSVERSIONINFO
        Dim dwOSVersionInfoSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
        'Dim szCSDVersion As String '*128
        <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, sizeconst:=128)> Public szCSDVeresion As String

    End Structure

    'Public Structure DEVMODE
    '    Dim dmDeviceName As String '* CCHDEVICENAME
    '    Dim dmSpecVersion As Integer
    '    Dim dmDriverVersion As Integer
    '    Dim dmSize As Integer
    '    Dim dmDriverExtra As Integer
    '    Dim dmFields As Long
    '    Dim dmOrientation As Integer
    '    Dim dmPaperSize As Integer
    '    Dim dmPaperLength As Integer
    '    Dim dmPaperWidth As Integer
    '    Dim dmScale As Integer
    '    Dim dmCopies As Integer
    '    Dim dmDefaultSource As Integer
    '    Dim dmPrintQuality As Integer
    '    Dim dmColor As Integer
    '    Dim dmDuplex As Integer
    '    Dim dmYResolution As Integer
    '    Dim dmTTOption As Integer
    '    Dim dmCollate As Integer
    '    Dim dmFormName As String '* CCHFORMNAME
    '    Dim dmLogPixels As Integer
    '    Dim dmBitsPerPel As Long
    '    Dim dmPelsWidth As Long
    '    Dim dmPelsHeight As Long
    '    Dim dmDisplayFlags As Long
    '    Dim dmDisplayFrequency As Long
    '    Dim dmICMMethod As Long        ' // Windows 95 only
    '    Dim dmICMIntent As Long        ' // Windows 95 only
    '    Dim dmMediaType As Long        ' // Windows 95 only
    '    Dim dmDitherType As Long       ' // Windows 95 only
    '    Dim dmReserved1 As Long        ' // Windows 95 only
    '    Dim dmReserved2 As Long        ' // Windows 95 only
    'End Structure

    Public Structure PRINTER_INFO_5
        Dim pPrinterName As String
        Dim pPortName As String
        Dim Attributes As Long
        Dim DeviceNotSelectedTimeout As Long
        Dim TransmissionRetryTimeout As Long
    End Structure

    Public Structure PRINTER_DEFAULTS
        Dim pDatatype As Long
        Dim pDevMode As Long
        Dim DesiredAccess As Long
    End Structure

    Declare Function GetProfileString Lib "kernel32" _
    Alias "GetProfileStringA" _
    (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long

    Declare Function WriteProfileString Lib "kernel32" _
    Alias "WriteProfileStringA" _
    (ByVal lpszSection As String, _
    ByVal lpszKeyName As String, _
    ByVal lpszString As String) As Long

    Declare Function SendMessage2 Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lparam As String) As Long

    Public Declare Function SetActiveWindow Lib "user32" Alias "SetActiveWindow" (ByVal hwnd As Int32) As Int32

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As IntPtr, _
            ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Public Declare Function GetWindowText Lib "user32" Alias _
        "GetWindowTextA" (ByVal hwnd As Int32, ByVal lpString As String, _
        ByVal cch As Int32) As Int32

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function FindWindow(ByVal className As String, ByVal ByValWindowsName As String) As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function FindWindowEx(ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr

    End Function

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function GetDlgCtrlID(ByVal hwnd As IntPtr) As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function GetDesktopWindow() As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("user32")> _
    Public Function GetWindow(ByVal hWnd As IntPtr, ByVal wCmd As Long) As IntPtr

    End Function


    Public Const WM_COMMAND = &H111
    Public Const WM_CLOSE = &H10
    Public Const WM_SETTEXT = &HC
    Public Const BM_SETSTATE = &HF3
    Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
    Public Const WM_LBUTTONDOWN = &H201     'Button down
    Public Const WM_LBUTTONUP = &H202       'Button up

    'Public Lectura As Boolean
    'Public Escritura As Boolean


    'Declare Function GetVersionExA Lib "kernel32" _
    '(ByVal lpVersionInformation As OSVERSIONINFO) As Integer
    Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" _
    (ByRef lpVersionInformation As OSVERSIONINFO) As Integer

    Public Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" _
    (ByVal pPrinterName As String, _
    ByVal phPrinter As Long, _
    ByVal pDefault As PRINTER_DEFAULTS) As Long

    Public Declare Function SetPrinter Lib "winspool.drv" _
    Alias "SetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    ByVal pPrinter As Object, _
    ByVal Command As Long) As Long    'Object-any

    Public Declare Function GetPrinter Lib "winspool.drv" _
    Alias "GetPrinterA" _
    (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    ByVal pPrinter As Object, _
    ByVal cbBuf As Long, _
    ByVal pcbNeeded As Long) As Long

    Public Declare Function lstrcpy Lib "kernel32" _
    Alias "lstrcpyA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As Object) As Long

    Public Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

    Public Sub SelectPrinter(ByVal NewPrinter As String, ByRef Printer As String)

        Dim strPrinter As String

        For Each strPrinter In Printing.PrinterSettings.InstalledPrinters
            If strPrinter = NewPrinter Then
                Printer = strPrinter
                Exit For
            End If
        Next

    End Sub

End Module

