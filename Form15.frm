VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form15 
   BorderStyle     =   0  'None
   Caption         =   "Form15"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form15.frx":0000
   ScaleHeight     =   4110
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   480
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar Pb 
      DragMode        =   1  'Automatic
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "loading"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   7770
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.Show
    admin (False)
End Sub

Private Sub Timer1_Timer()
    Pb.Value = Pb.Value + 1
    If Pb.Value = 100 Then Unload Me
    Label1.Caption = Pb.Value & "% loading.."
End Sub
