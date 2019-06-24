VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   Caption         =   "Update Keaktifan"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   4710
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton button1 
      Default         =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "UPDATE"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   6615
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3960
         TabIndex        =   0
         Top             =   200
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Keaktifan"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   6615
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tahun"
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Ikhwah"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   6615
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label7 
         Caption         =   "Tempat Tarbiyah"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Nama KTI"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Murobbi"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Ikhwah"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6615
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "No Telp"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Nama"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin DBLiqo.jcbutton button2 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "BATAL"
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiCombo()
    With Combo1
        .Clear
        .AddItem "Januari"
        .AddItem "Februari"
        .AddItem "Maret"
        .AddItem "April"
        .AddItem "Mei"
        .AddItem "Juni"
        .AddItem "Juli"
        .AddItem "Agustus"
        .AddItem "September"
        .AddItem "Oktober"
        .AddItem "November"
        .AddItem "Desember"
    End With
    
    With Combo2
        .Clear
        .AddItem "Aktif"
        .AddItem "Tidak Aktif"
        .AddItem "Kurang"
    End With
End Sub

Private Sub button1_Click()
Dim Sql As String
    If Combo1.text <> "" Then
        With rsupIk
            Sql = "NamaL='" & lbl(0).Caption & "' AND wkt='" & Text1.text & "-" & Combo1.ListIndex + 1 & "'"
            .Filter = Sql
            If .EOF Then .AddNew
            
            !namaL = lbl(0).Caption
            !Hp = lbl(1).Caption
            !KKI = lbl(2).Caption
            !wkt = Text1.text & "-" & Combo1.ListIndex + 1
            !keaktifan = Combo2.text
            .Update
            .Requery
            
            Unload Me
        End With
    End If
End Sub

Private Sub button2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    isiCombo
    Combo1.ListIndex = Format(Date, "m") - 1
    Text1.text = Format(Date, "YYYY")
End Sub

