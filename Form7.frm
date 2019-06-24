VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Input Data Ustadz"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton Btn2 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "BATAL"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton Btn1 
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SIMPAN"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txt 
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   4935
   End
   Begin VB.PictureBox XPFrame2 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      Begin VB.Label Label2 
         Caption         =   "ALAMAT"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "NO TELP/HP"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "NAMA"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox XPFrame1 
      Height          =   2655
      Left            =   1800
      ScaleHeight     =   2595
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txt 
         Height          =   975
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txt 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn1_Click()
    If txt(0).text <> "" Or txt(1).text <> "" Or txt(2).text <> "" Then
        
        With rsUst
        .Find "Nama='" & "Ust. " & txt(0).text & "'"
        If .EOF Then
            If Baru Then .AddNew
            
            !Nama = txt(0).text
            !Alamat = txt(1).text
            !Hp = txt(2).text
            .Update
            .Requery
            'MsgBox "Data telah tersimpan..!", vbInformation, "Tersimpan"
            Unload Me
        Else
            MsgBox "Data ini sudah ada..."
            Unload Me
        End If
        End With
        
    Else
        MsgBox "Penginputan kurang lengkap..!", vbApplicationModal, "Maaf"
    End If
End Sub

Private Sub XPButton1_Click()
End Sub

Private Sub Btn2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub
