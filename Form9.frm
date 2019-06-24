VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Detail Data Ustadz Murobbi"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton XPbutton3 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
      Caption         =   "X"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPbutton2 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Caption         =   ">>"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPbutton1 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Caption         =   "<<"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.PictureBox XPFrame1 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   8715
      TabIndex        =   3
      Top             =   2040
      Width           =   8775
      Begin DBLiqo.jcbutton XPButton4 
         Height          =   615
         Left            =   7320
         TabIndex        =   12
         Top             =   2880
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
         Caption         =   "CETAK"
         UseMaskCOlor    =   -1  'True
      End
      Begin Crystal.CrystalReport CrUstD 
         Left            =   480
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form9.frx":26E33
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Double Click Untuk melihat Data KKI Secara Lengkap"
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   20
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Data KKI"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "NamaKKI"
            Caption         =   "NAMA KKI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Tahun"
            Caption         =   "TAHUN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Murobbi"
            Caption         =   "MUROBBI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Naqib"
            Caption         =   "NAQIB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Tempat"
            Caption         =   "TEMPAT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Hari"
            Caption         =   "HARI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Waktu"
            Caption         =   "WAKTU"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Kecamatan"
            Caption         =   "KECAMATAN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Jenis"
            Caption         =   "TINGKAT"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JUMLAH HALAQAH YANG DI TANGANI"
         ForeColor       =   &H00D54600&
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   3000
         Width           =   2925
      End
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00D54600&
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00D54600&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00D54600&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NO TELP"
      ForeColor       =   &H00D54600&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT LENGKAP"
      ForeColor       =   &H00D54600&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA USTADZ"
      ForeColor       =   &H00D54600&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ValData()
If Not rsUst.EOF Then
    lbl(0).Caption = rsUst!nama
    lbl(1).Caption = rsUst!Alamat
    lbl(2).Caption = rsUst!Hp
    rsKKI.Filter = "Murobbi='" & rsUst!nama & "'"
    Label4.Caption = "JUMLAH HALAQAH YANG DITANGANI    : " & rsKKI.RecordCount & " Liqo"
End If

End Sub

Private Sub DataGrid1_DblClick()
    Form5.Show
End Sub

Private Sub Form_Load()
    ValData
    Set DataGrid1.DataSource = rsKKI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsKKI.Filter = ""
End Sub

Private Sub XPButton1_Click()
If Not rsUst.EOF Then
    If Not rsUst.AbsolutePosition <= 1 Then rsUst.MovePrevious
        ValData
End If
End Sub

Private Sub XPbutton2_Click()
If Not rsUst.EOF Then
    If Not rsUst.AbsolutePosition >= rsUst.RecordCount Then rsUst.MoveNext
    ValData
End If
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub

Private Sub XPButton4_Click()
    With CrUstD
        .ReportFileName = App.Path & "\RPT\RptUstD.rpt"
        .DataFiles(0) = App.Path & "\Db.mdb"
        .WindowState = crptMaximized
        .ReplaceSelectionFormula "{DBUst.Nama}='" & lbl(0).Caption & "'"
        .Action = 1
        .Reset
    End With
End Sub
