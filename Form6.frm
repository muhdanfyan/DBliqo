VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Biodata Ikhwah"
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   13125
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton BtnKlr 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   11880
      TabIndex        =   46
      Top             =   6360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      Caption         =   "TUTUP"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton CmCtk 
      Height          =   495
      Left            =   7440
      TabIndex        =   45
      Top             =   6240
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
      Caption         =   "CETAK"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton1 
      Height          =   615
      Left            =   360
      TabIndex        =   44
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "<<"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton2 
      Height          =   615
      Left            =   1320
      TabIndex        =   43
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   ">>"
      UseMaskCOlor    =   -1  'True
   End
   Begin Crystal.CrystalReport CrBio 
      Left            =   9000
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin DBLiqo.jcbutton btnSMS 
      Height          =   495
      Left            =   7440
      TabIndex        =   47
      Top             =   6720
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
      Caption         =   "SMS IKHWAH"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "GOLONGAN DARAH"
      Height          =   255
      Left            =   7560
      TabIndex        =   42
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "NO TELP DAN HP"
      Height          =   255
      Left            =   7560
      TabIndex        =   41
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   19
      Left            =   9600
      TabIndex        =   40
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   18
      Left            =   9600
      TabIndex        =   39
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   17
      Left            =   9600
      TabIndex        =   38
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   16
      Left            =   9600
      TabIndex        =   37
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   15
      Left            =   9480
      TabIndex        =   36
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   14
      Left            =   9480
      TabIndex        =   35
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   13
      Left            =   9600
      TabIndex        =   34
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   9600
      TabIndex        =   33
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   11
      Left            =   2640
      TabIndex        =   32
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   31
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   30
      Top             =   5520
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   29
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   28
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   27
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   26
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   25
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   24
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   23
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7800
      TabIndex        =   22
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7800
      TabIndex        =   21
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "IBU"
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "AYAH"
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PEKERJAAN ORANG TUA"
      Height          =   255
      Left            =   7440
      TabIndex        =   18
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPAT/HARI/WAKTU"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "MURABBI"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA HALAQAH"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MARHALAH"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "PENGALAMAN ORGANISASI"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "3) SMA"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "2) SMP"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "1) SD"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "RIWAYAT PENDIDIKAN"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH BERSAUDARA"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "BAKAT\ HOBBY"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PT\FAK\JUR\ANGKATAN"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "2. ALAMAT MAKASSAR"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1. ALAMAT DAERAH"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPAT/TANGGAL LAHIR"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA KUNNIYAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA LENGKAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ValData()
Dim i As Integer
If Not rsIkh.EOF Or Not rsIkh.BOF Then
    lbl(0).Caption = rsIkh!namaL
    lbl(1).Caption = rsIkh!NamaKunn
    lbl(2).Caption = rsIkh!Tempat & ", " & rsIkh!tanggal
    lbl(3).Caption = rsIkh!AlamatD
    lbl(4).Caption = rsIkh!AlamatM
    lbl(5).Caption = rsIkh!NamaS & "\" & rsIkh!Fak & "\" & rsIkh!Angkt
    lbl(6).Caption = rsIkh!Bkt
    lbl(7).Caption = "Anak Ke- " & rsIkh!Usaud & " Dari " & rsIkh!Jsaud & " Bersaudara"
    lbl(8).Caption = rsIkh!RSD
    lbl(9).Caption = rsIkh!RSMP
    lbl(10).Caption = rsIkh!RSMA
    lbl(11).Caption = rsIkh!Porg
    lbl(12).Caption = rsIkh!Tingkat
    lbl(13).Caption = rsIkh!NmHalaqa
    lbl(16).Caption = rsIkh!Ayah
    lbl(17).Caption = rsIkh!Ibu
    lbl(18).Caption = rsIkh!Hp
    lbl(19).Caption = rsIkh!Gol
    rsKKI.Find "NamaKKI='" & lbl(13).Caption & "'"
    If Not rsKKI.EOF Then
        lbl(14).Caption = rsKKI!murobbi
        lbl(15).Caption = rsKKI!Tempat & " \ " & rsKKI!Hari & " \ " & rsKKI!Waktu & " WITA"
    Else
        lbl(14).Caption = ""
        lbl(15).Caption = "- \ - \ -"
    End If
    rsKKI.Requery
Else
    For i = 0 To lbl.Count - 1
        lbl(i).Caption = "-"
    Next i
End If
End Sub

Private Sub BtnKlr_Click()
    Unload Me
End Sub

Private Sub btnSMS_Click()
    If Not rsIkh.EOF Then
        With Form14
            .Tab1.Tabs(3).Selected = True
            .List1.AddItem rsIkh!Hp
            .Show
        End With
    End If
End Sub

Private Sub CmCtk_Click()
    With CrBio
        .ReportFileName = App.Path & "\RPT\RptIkhD.rpt"
        .DataFiles(0) = App.Path & "\DB.mdb"
        .ReplaceSelectionFormula "{DBIkhwah.NamaL}='" & lbl(0).Caption & "'"
        .WindowState = crptMaximized
        .Action = 1
        .Reset
    End With
End Sub

Private Sub CmTk_Click()

End Sub

Private Sub Form_Load()
    ValData
End Sub

Private Sub jcbutton1_Click()

End Sub

Private Sub XPButton1_Click()
If Not rsIkh.EOF Then
    If Not rsIkh.AbsolutePosition = 1 Then rsIkh.MovePrevious
    ValData
End If
End Sub

Private Sub XPbutton2_Click()
If Not rsIkh.EOF Then
    If Not rsIkh.AbsolutePosition = rsIkh.RecordCount Then rsIkh.MoveNext
    ValData
End If
End Sub
