VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "f"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   4935
      Begin VB.ComboBox CbJen 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   4695
         Begin VB.ComboBox cb 
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   3360
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   525
            Index           =   6
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   9
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label9 
            Caption         =   "Kecamatan"
            ForeColor       =   &H00D54600&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Tempat"
            ForeColor       =   &H00D54600&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Hari"
            ForeColor       =   &H00D54600&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Waktu"
            ForeColor       =   &H00D54600&
            Height          =   255
            Left            =   2280
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox Cbt 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   4440
         Width           =   2655
      End
      Begin VB.ComboBox CbBd 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2400
         Width           =   3015
      End
      Begin VB.ComboBox CbNq 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ComboBox Cbm 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   405
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label11 
         Caption         =   "Jenis"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Tingkat"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Bendahara"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Naqib"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Murabbi"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tahun Daurah"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nama KKI"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin DBLiqo.jcbutton button1 
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   6
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
      BackColor       =   16744576
      Caption         =   "Simpan"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button2 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   6
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
      BackColor       =   16744576
      Caption         =   "Batal"
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiCombo()
    
    With Cbm
    If Not rsUst.EOF Then
        rsUst.MoveFirst
        Do While Not rsUst.EOF
            .AddItem "Ust. " & rsUst!Nama
            rsUst.MoveNext
        Loop
        rsUst.MoveFirst
    End If
    End With
    With cb
        .Clear
        .AddItem "Senin"
        .AddItem "Selasa"
        .AddItem "Rabu"
        .AddItem "Kamis"
        .AddItem "Jum'at"
        .AddItem "Sabtu"
        .AddItem "Minggu"
    End With
    
    With Cbt
        .Clear
        .AddItem "PRA TAARIF"
        .AddItem "TAARIFIYAH"
        .AddItem "TAKWINIYAH"
    End With
    
    With CbJen
        .Clear
        .AddItem "Sekolah"
        .AddItem "Kampus"
        .AddItem "Umum"
    End With
End Sub

Private Sub button1_Click()
If Text1(0).text <> "" Or Text1(1).text <> "" Or CbNq.text <> "" Or CbBd.text <> "" Or Text1(5).text <> "" Or Text1(6).text <> "" Or Text1(7).text <> "" Or cb.text <> "" Or Cbt.text <> "" Then
    If Baru Then
        rsKKI.AddNew
    End If
    
    With rsKKI
        !namaKKI = Text1(0).text
        !Tahun = Text1(1).text
        If Cbm.text = "" Then
            !murobbi = "-"
        Else
        !murobbi = Right(Cbm.text, Len(Cbm.text) - 5)
        End If
        !Naqib = CbNq.text
        !Bendahara = CbBd.text
        !Hari = cb.text
        !Waktu = Text1(5).text
        !Tempat = Text1(6).text
        !Kecamatan = Text1(7).text
        !Jenis = Cbt.text
        !Tipe = CbJen.text
        .Update
        .Requery
    End With
'    MsgBox "Data Telah tersimpan..!", vbInformation, "Ok"
    Unload Me
Else
    MsgBox "Penginputan data masih belum lengkap..!", vbApplicationModal, "Maaf!"
End If
End Sub

Private Sub button2_Click()
    Unload Me
End Sub

Private Sub Cbm_Change()
Dim a As String
End Sub

Private Sub Form_Load()
    isiCombo
End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 0 Then
        rsKKI.Find "NamaKKI='" & Text1(0).text & "'"
        If Not rsKKI.EOF Then
            Baru = False
            Text1(0).text = rsKKI!namaKKI
            Text1(1).text = rsKKI!Tahun
            Cbm.text = "Ust. " & rsKKI!murobbi
            CbJen.text = rsKKI!Tipe
            CbNq.text = rsKKI!Naqib
            CbBd.text = rsKKI!Bendahara
            cb.text = rsKKI!Hari
            Text1(5).text = rsKKI!Waktu
            Text1(6).text = rsKKI!Tempat
            Text1(7).text = rsKKI!Kecamatan
            Cbt.text = rsKKI!Jenis
            CbJen.text = rsKKI!Tipe
        End If

    End If
End Sub
