Attribute VB_Name = "Module2"
Option Explicit

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const WM_LBUTTONDBLCLICK = &H203 'left button click
Public Const WM_LBUTTONDOWN = &H201     'left button pressed
Public Const WM_LBUTTONUP = &H202       'left button lifted

Public Const WM_RBUTTONDBLCLK = &H206   'right button click
Public Const WM_RBUTTONDOWN = &H204     'right button pressed
Public Const WM_RBUTTONUP = &H205       'right button lifted

Public Const WM_MOUSEMOVE = &H200       'mouse moved

Public Const NIM_ADD = &H0              'icon is being added
Public Const NIM_MODIFY = &H1           'icon has been modified
Public Const NIM_DELETE = &H2           'icon has been removed

Public Const NIF_MESSAGE = &H1          'windows message sent
Public Const NIF_ICON = &H2             'an icon
Public Const NIF_INFO = &H10            '0_o
Public Const NIF_TIP = &H4              'o_0
     
Public Const NIS_SHAREDICON = &H2       '0_o

Public Type NOTIFYICONDATA              'This is the data type you must use for tray icons
   cbSize As Long                       'size of this
   hWnd As Long                         'the windows handle on this 'window'
   uID As Long                          'dunno
   uFlags As Long                       'flags tell the computer some stuff about the icon
   uCallbackMessage As Long             'what is returned
   hIcon As Long                        'the picture for this icon
   szTip As String * 128                'the tool tip text (shown when moouse is hovered over icon)
   dwState As Long                      ' not 100% sure yet
   dwStateMask As Long                  '"                 "
   szInfo As String * 256               '"                 "
   uTimeout As Long                     'time before timeout (i believe)
   szInfoTitle As String * 64           '0_o
   dwInfoFlags As Long                  'o_0
End Type                                'I believe some of these can be left out, but incorrectly structured types can seriously crash windows!

Public TrayIcon As NOTIFYICONDATA

Public Sub AddTray(Frm As Form)
     With TrayIcon
          .cbSize = Len(TrayIcon)                 'just do it
          .uID = vbNull                           'just do it
          .hWnd = Frm.hWnd                        'the hWnd of the form this will 'come from'
          .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO    '0_o  just do it
          .uCallbackMessage = WM_MOUSEMOVE        'what things will activate a sub in frm - this activates frm_MouseMove when the mouse goes over the tray icon
          .dwState = NIS_SHAREDICON               'just do it
          .hIcon = Frm.Icon                       'the picture of the tray icon
          .szTip = "Backup Progam" & vbNullChar   'the ToolTipText
     End With
     Shell_NotifyIcon NIM_ADD, TrayIcon           'tell the tray we're adding an Icon
     Frm.Hide                                    'uncomment to autohide the calling form
End Sub
