VERSION 5.00
Begin VB.Form FLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form16"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FLogin.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin DBLiqo.jcbutton jcbutton1 
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Login"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin DBLiqo.jcbutton jcbutton2 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Batal"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = Screen.Height - Me.Height - 1300
    Me.Left = Screen.Width - FLogin.Width - 100
    refreshData
End Sub

Private Sub jcbutton1_Click()

user.Filter = "user='" & Text1.text & "' and password='" & Text2.text & "'"
If user.EOF Then
    MsgBox "User dan password tidak valid...", vbCritical, "Maaf!"
    Text1.text = ""
    Text2.text = ""
    Text1.SetFocus
Else
    FormUtama.Show
    admin (True)
    Unload Me
End If

End Sub

Private Sub jcbutton2_Click()
    Unload Me
End Sub
