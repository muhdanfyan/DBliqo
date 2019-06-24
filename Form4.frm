VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "INPUT DATA IKHWAH"
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   11280
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton XPButton2 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   1680
      TabIndex        =   58
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "BATAL"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton1 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   57
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "SIMPAN"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.PictureBox XPFrame9 
      Height          =   3855
      Left            =   6120
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   48
      Top             =   5160
      Width           =   5055
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   3855
      End
      Begin VB.ComboBox CbGol 
         Height          =   315
         Left            =   2880
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   13
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   12
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "KEAHLIAN"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label21 
         Caption         =   "NOMOR YANG BISA DIHUBUNGI"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "PENGALAMAN ORGANISASI"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "GOLONGAN DARAH"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   1080
         TabIndex        =   49
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.PictureBox XPFrame7 
      Height          =   975
      Left            =   3840
      ScaleHeight     =   915
      ScaleWidth      =   7275
      TabIndex        =   42
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   0
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "NAMA LENGKAP"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "NAMA KUNNIYAH"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox XPFrame6 
      Height          =   1215
      Left            =   6120
      ScaleHeight     =   1155
      ScaleWidth      =   4995
      TabIndex        =   39
      Top             =   3840
      Width           =   5055
      Begin VB.ComboBox CbH1 
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblT 
         ForeColor       =   &H00D54600&
         Height          =   315
         Left            =   1320
         TabIndex        =   56
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label16 
         Caption         =   "TINGKAT"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "NAMA"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox XPFrame5 
      Height          =   1575
      Left            =   6120
      ScaleHeight     =   1515
      ScaleWidth      =   4995
      TabIndex        =   35
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   1200
         TabIndex        =   16
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label14 
         Caption         =   "SMA"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "SMP"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "SD"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox XPFrame4 
      Height          =   855
      Left            =   6120
      ScaleHeight     =   795
      ScaleWidth      =   4995
      TabIndex        =   31
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "BERSAUDARA"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "DARI"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "ANAK KE"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox XPFrame3 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   5835
      TabIndex        =   27
      Top             =   4320
      Width           =   5895
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   1920
         TabIndex        =   9
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   3855
      End
      Begin VB.ComboBox CbPen 
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "JURUSAN"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "ANGKATAN"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "NAMA"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "FAKULTAS"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.PictureBox XPFrame2 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   5835
      TabIndex        =   24
      Top             =   2280
      Width           =   5895
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "ALAMAT MAKASSAR"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "ALAMAT DAERAH"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox XPFrame1 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   5835
      TabIndex        =   23
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CalendarForeColor=   13977088
         CalendarTitleForeColor=   13977088
         CalendarTrailingForeColor=   13977088
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   99942403
         CurrentDate     =   40218
      End
      Begin VB.Label Label22 
         Caption         =   "TANGGAL"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "TEMPAT"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox XPFrame8 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   5835
      TabIndex        =   45
      Top             =   6840
      Width           =   5895
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   1320
         TabIndex        =   11
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label20 
         Caption         =   "AYAH"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "IBU"
         ForeColor       =   &H00D54600&
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiCombo()
    With CbPen
        .Clear
        .AddItem "SD"
        .AddItem "SMP"
        .AddItem "SMA"
        .AddItem "KULIAH"
        .AddItem "Tidak Ada"
    End With
    
    
    
    With CbH1
        .Clear
        If rsKKI.PageCount > 0 Then
            rsKKI.MoveFirst
            Do While Not rsKKI.EOF
                .AddItem rsKKI!namaKKI
                rsKKI.MoveNext
            Loop
            rsKKI.MoveFirst
        End If
        
    End With
    
    With CbGol
        .Clear
        .AddItem "A"
        .AddItem "B"
        .AddItem "AB"
        .AddItem "O"
        .text = "-"
    End With
    
End Sub
Sub bersih()
    Dim i As Integer
    
    For i = 0 To txt.Count - 1
        txt(i).text = ""
    Next i
    
End Sub

Sub Pend()
If CbPen.text = "Tidak Ada" Then
    txt(17).text = "-"
    txt(5).text = "-"
    txt(6).text = "-"
    txt(18).text = "-"
    txt(18).Enabled = False
    txt(17).Enabled = False
    txt(5).Enabled = False
    txt(6).Enabled = False
Else
    txt(18).Enabled = True
    txt(17).Enabled = True
    txt(5).Enabled = True
    txt(6).Enabled = True
    txt(17).text = ""
    txt(5).text = ""
    txt(6).text = ""
    If CbPen.text = "SMP" Then
        txt(10).Enabled = True
        txt(10).text = ""
        txt(11).text = "-"
        txt(11).Enabled = False
    ElseIf CbPen.text = "SD" Then
        txt(10).text = "-"
        txt(10).Enabled = False
        txt(11).text = "-"
        txt(11).Enabled = False
    Else
        txt(10).Enabled = True
        txt(11).Enabled = True
        txt(10).text = ""
        txt(11).text = ""
    End If
End If
End Sub

Private Sub CbGol_Change()
If CbGol.ListIndex >= 0 Then XPButton1.Default = True
End Sub

Private Sub Cbh1_Click()
    If CbH1.ListIndex >= 0 Then
        rsKKI.Filter = "NamaKKI='" & CbH1.text & "'"
        lblT.Caption = rsKKI!Jenis
    Else
        rsKKI.Filter = ""
    End If
End Sub

Private Sub CbPen_Change()
    Pend
End Sub

Private Sub CbPen_Click()
    Pend
End Sub

Private Sub Form_Load()
    isiCombo
'    refreshData
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 Then
        rsIkh.Find "NamaL='" & txt(0).text & "'"
        If Not rsIkh.EOF Then
        Baru = False
            txt(0).text = rsIkh!namaL
            txt(1).text = rsIkh!NamaKunn
            txt(2).text = rsIkh!Tempat
            If rsIkh!tanggal <> Empty Then DT1.Value = rsIkh!tanggal
            txt(3).text = rsIkh!AlamatD
            txt(4).text = rsIkh!AlamatM
            CbPen.text = rsIkh!Pendidikan
            txt(17).text = rsIkh!NamaS
            txt(18).text = rsIkh!jur
            txt(5).text = rsIkh!Fak
            txt(6).text = rsIkh!Angkt
            txt(7).text = rsIkh!Usaud
            txt(8).text = rsIkh!Jsaud
            txt(9).text = rsIkh!RSD
            txt(10).text = rsIkh!RSMP
            txt(11).text = rsIkh!RSMA
            txt(12).text = rsIkh!Bkt
            txt(13).text = rsIkh!Porg
            txt(14).text = rsIkh!Ayah
            txt(15).text = rsIkh!Ibu
            txt(16).text = rsIkh!Hp
            CbH1.text = rsIkh!NmHalaqa
            lblT.Caption = rsIkh!Tingkat
        End If
    End If
End Sub

Private Sub XPButton1_Click()
Dim namaL, NamaKunn, Tempat, tanggal, AlamatD, AlamatM, Pendidikan, NamaS, Fak, jur, Angkt, Bkt, Usaud, Jsaud, RSD, RSMP, RSMA, Porg, _
Tingkat, NmHalaqa, Ayah, Ibu, Hp, Gol As String
Dim sblm As Boolean
sblm = False
If txt(0).text <> "" Or txt(1).text <> "" Or txt(2).text <> "" Or txt(3).text <> "" Or txt(4).text <> "" Or txt(16).text <> "" Or CbH1.text <> "" Then
    If Baru Then
        rsIkh.AddNew
    End If
        With rsIkh
            !namaL = txt(0).text
            !NamaKunn = txt(1).text
            !Tempat = txt(2).text
            !tanggal = DT1.Value
            !AlamatD = txt(3).text
            !AlamatM = txt(4).text
            !Pendidikan = CbPen.text
            !NamaS = txt(17).text
            !Fak = txt(5).text
            !jur = txt(18).text
            !Angkt = txt(6).text
            !Bkt = txt(12).text
            !Usaud = txt(7).text
            !Jsaud = txt(8).text
            !RSD = txt(9).text
            !RSMP = txt(10).text
            !RSMA = txt(11).text
            !Porg = txt(13).text
            !Tingkat = lblT.Caption
            !NmHalaqa = CbH1.text
            !Ayah = txt(14).text
            !Ibu = txt(15).text
            !Hp = txt(16).text
            !Gol = CbGol.text
            .Update
            .Requery
            rsKKI.Requery
        End With
'        MsgBox "Data Telah tersimpan...", vbInformation, "Tersimpan!"
        
Else
    MsgBox "Data Penginputan tidak lengkap...!", vbApplicationModal, "Maaf..!"
End If
    Unload Me
End Sub

Private Sub XPbutton2_Click()
    Unload Me
End Sub
