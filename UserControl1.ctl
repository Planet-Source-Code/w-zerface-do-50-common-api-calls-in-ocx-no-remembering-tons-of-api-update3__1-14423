VERSION 5.00
Begin VB.UserControl CompControl 
   BackColor       =   &H000000FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   495
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "UserControl1.ctx":0442
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "CompControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'By Martin McCormick
Dim a123
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Const Internet_Autodial_Force_Unattended As Long = 2
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim retval
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private XT(1) As Single, YT As Single, M As Single
Private XScreen As Single, YScreen As Single
Private i As Integer, II As Integer
Private Const MAX_DELOCATION = 250
Private Const PUPIL_DISTANCE = 30
Option Explicit
Dim timeval
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
Private Const SW_SHOW = 5
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Dim dgf
Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4


Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim nid As NOTIFYICONDATA
Private Declare Function SwapMouseButton Lib "user32.dll" (ByVal bSwap As Long) As Long
'*************************************************************
'*************************************************************
'*************************************************************
'NEW API!:
'KERNAL32:
Private Declare Function Beep _
    Lib "kernel32" ( _
        ByVal dwFreq As Long, _
        ByVal dwDuration As Long) _
    As Long
Private Declare Function CloseHandle _
    Lib "kernel32" ( _
        ByVal hObject As Long) _
    As Long
Private Declare Function GetFileSize _
    Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpFileSizeHigh As Long) _
    As Long
Private Declare Function GetLastError _
    Lib "kernel32" ( _
        ) _
    As Long
Private Declare Function GetVersion _
    Lib "kernel32" () _
    As Long
Private Declare Function SetCurrentDirectory _
    Lib "kernel32" _
    Alias "SetCurrentDirectoryA" ( _
        ByVal lpPathName As String) _
    As Long
Private Declare Function GetPixel _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long) _
    As Long
Private Declare Function SetPixel _
    Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) _
    As Long

'*************************************************************
'*************************************************************
'*************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''
'Event Declarations:
Event ShutDown(ShutDown)
Event Restart(Restart)
Event LogOff(LogOff)
Event TaskBarHide(TaskBarHide)
Event TasksBarShow(TasksBarShow)
Event ScreenSaverOn(ScreenSaverOn)
Event ScreenSaverOff(ScreenSaverOff)
Event DesktopIconsHide(DesktopIconsHide)
Event ALTCTRLDELEnabled(ALT_CTRL_DEL_Enabled)
Event ALTCTRLDELDisabled(ALT_CTRL_DEL_Disabled)
Event OpenCDROM(OpenCDROM)
Event EmptRecycle(EmptRecycle)
Event MinimizeAll(MinimizeAll)
Event OpenExplore(OpenExplore)
Event FindFiles(FindFiles)
Event OpenInternetBrowser(OpenInternetBrowser)
Event InternetConnect(InternetConnect)
Event InternetDiconnect(InternetDiconnect)
Event SendEmail(SendEmail)
Event AddRemove(Add_Remove)
Event AddHardWare(Add_HardWare)
Event TimeDateSettings(Time_Date_Settings)
Event RegionalSettings(Regional_Settings)
Event DisplaySettings(Display_Settings)
Event InternetSetting(Internet_Settings)
Event KeyboardSettings(Keyboard_Settings)
Event MouseSettings(Mouse_Settings)
Event ModemSettings(Modem_Settings)
Event SystemSettings(System_Settings)
Event NetworkSettings(Network_Settings)
Event PasswordSettings(Password_Settings)
Event SoundsSettings(Sounds_Settings)
Event ShowAbout(ShowAbout)
Event CopyaFile(Copy_File)
Event DeleteaFile(Delete_File)
Event MoveaFile(Move_File)
Event FlipMouseButtons(FlipMouseButtons)
Event FormToTop(FormOnTop)
Event SpecialBeep(Special_beep)
Event CloseObjectHandle(Close_Object_Handle)
Event LastError(Last_error)
Event WinVersion(Win_Version)
Event SetCurDir(Set_Cur_Dir)
Event PixelSet(Pixel_Set)
Event PixelGet(Pixel_Get)

Private Sub UserControl_Resize()
UserControl.Width = 500
UserControl.Height = 500
End Sub
Function ShutDown()
ShutDown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function
Function Restart()
Restart = ExitWindowsEx(EWX_REBOOT, 0&)
End Function
Function LogOff()
LogOff = ExitWindowsEx(EWX_LOGOFF, 0&)
End Function
Function TaskBarHide()
Dim rtn
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Function
Function TaskBarShow()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Function
Function ScreenSaverOn()
ToggleScreenSaverActive (True)
End Function
Function ScreenSaverOff()
ToggleScreenSaverActive (False)
End Function
Public Function ToggleScreenSaverActive(Active As Boolean) _
   As Boolean
Dim lActiveFlag As Long
Dim retval As Long

lActiveFlag = IIf(Active, 1, 0)
retval = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, _
   lActiveFlag, 0, 0)
ToggleScreenSaverActive = retval > 0

End Function
Function DesktopIconsShow()
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 5
End Function
Function DesktopIconsHide()
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 0
End Function
Function ALT_CTRL_DEL_Enabled()
callme (False)
End Function
Function ALT_CTRL_DEL_Disabled()
callme (True)
End Function
Private Sub callme(huh As Boolean)
Dim gd
gd = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub
Function OpenCDROM()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Function
Function EmptRecycle()
EmptRecycle = SHEmptyRecycleBin(UserControl.hWnd, "", SHERB_NOPROGRESSUI)
End Function
Function MinimizeAll()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function OpenExplore()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(69, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function FindFiles()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(70, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
Function OpenInternetBrowser()
ShellExecute hWnd, "open", "", vbNullString, vbNullString, conSwNormal
End Function
Function InternetConnect()
InternetConnect = InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
End Function
Function InternetDiconnect()
InternetDiconnect = InternetAutodialHangup(0&)
End Function
Function SendEmail()
ShellExecute hWnd, "open", "mailto:", vbNullString, vbNullString, SW_SHOW
End Function
Function Add_Remove()
Add_Remove = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Function

Function Add_HardWare()
Add_HardWare = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Function

Function Time_Date_Settings()
Time_Date_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Function

Function Regional_Settings()
Regional_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Function

Function Display_Settings()
Display_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Function

Function Internet_Settings()
Internet_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Function

Function Keyboard_Settings()
Keyboard_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Function

Function Mouse_Settings()
Mouse_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Function
Function Modem_Settings()
Modem_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Function
Function System_Settings()
System_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Function
Function Network_Settings()
Network_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Function
Function Password_Settings()
Password_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Function
Function Sounds_Settings()
Sounds_Settings = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Function
Function ShowAbout()
Dim about
about = MsgBox("CompControl.ocx was created by Martin McCormick and can be download from Http://www.planet-source-code.com any questions or comments should be sent to Slimshady_5_5_5@hotmail.com", vbOKOnly + vbInformation, "About")

End Function
Function Copy_File(FileToCopy, Destination)
Copy_File = CopyFile(FileToCopy, Destination, 1)
End Function
Function Delete_File(file)
Delete_File = DeleteFile(file)
End Function
Function Move_File(FileToMove, Destination)
Move_File = MoveFile(FileToMove, Destination)
End Function
Function FlipMouseButtons()
FlipMouseButtons = SwapMouseButton(1)
End Function
Function FormOnTop(Form, x, y, Width, Height)
Dim hnd
Width = Width / 15
Height = Height / 15
x = x / 15
y = y / 15
hnd = Form.hWnd
SetWindowPos hnd, conHwndTopmost, x, y, Width, Height, conSwpNoActivate Or conSwpShowWindow
End Function
Function Special_beep(Frequency, Duration)
Special_beep = Beep(Frequency, Duration)
End Function
Function Close_Object_Handle(Object)
Close_Object_Handle = CloseHandle(Object)
End Function
Function Last_error()
Last_error = GetLastError
End Function
Function Win_Version()
Win_Version = GetVersion
End Function
Function Set_Cur_Dir(Path)
Set_Cur_Dir = SetCurrentDirectory(Path)
End Function
Function Pixel_Set(hDC_Of_Object, X_Position_Of_Pixil, Y_Position_Of_Pixil, Color_To_Make_Pixil)
Pixel_Set = SetPixel(hDC_Of_Object, X_Position_Of_Pixil, Y_Position_Of_Pixil, Color_To_Make_Pixil)
End Function
Function Pixel_Get(hDC_Of_Object, X_Position_Of_Pixil, Y_Position_Of_Pixil)
Pixel_Get = GetPixel(hDC_Of_Object, X_Position_Of_Pixil, Y_Position_Of_Pixil)
End Function
