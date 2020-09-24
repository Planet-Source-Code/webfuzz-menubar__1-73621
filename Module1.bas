Attribute VB_Name = "Module1"

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const Twip_m As Single = 56692.854479
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public programtitle As String
Public MAX As Integer






Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type


Public Declare Function GetTickCount Lib "kernel32" () As Long






Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Public Const EWX_REBOOT = 2
Public Const EWX_LOGOFF = 0
Public Const EWX_FORCE = 4
Public Const EWX_SHUTDOWN = 1

Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long


Public Function Pause(Value As Long)  ' Value is the amount of time to pause for in milliseconds

Dim PreTick As Long

PreTick = GetTickCount

Do
DoEvents
If GetTickCount = PreTick + Value Then Exit Do
Loop

End Function







Function StartDoc(DocName As String) As Long
  Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
 
  'change "Open" to "Explore" to bring up file explorer
  StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", 1)
End Function





Function StartDoc2(DocName As String, parm As String) As Long
  Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
 
  'change "Open" to "Explore" to bring up file explorer
  StartDoc2 = ShellExecute(Scr_hDC, "Open", DocName, parm, "c:\", 1)
End Function




Sub logit(a As String)
On Error Resume Next
Open App.Path + "\log.txt" For Append As #2
Print #2, Str(Date) + " " + Str(Time) + " - " + a ' + vbCrLf
done:
Close #2
End Sub



