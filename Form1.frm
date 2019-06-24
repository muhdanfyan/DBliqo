VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DATA HALAQAH TARBIYAH"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton btAktif 
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Keaktifan KTI"
   End
   Begin VB.Frame Frame1 
      Caption         =   "SMS KKI"
      Height          =   2295
      Left            =   3530
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "Naqib dan bendahara"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Semua Ikhwah"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "SMS USTADZ"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "SMS BENDAHARA"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1700
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SMS NAQIB"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   1700
      End
      Begin DBLiqo.jcbutton button2 
         Height          =   375
         Left            =   2250
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         BackColor       =   14800597
         Caption         =   "TUTUP"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton button3 
         Height          =   375
         Left            =   2250
         TabIndex        =   17
         Top             =   1550
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         BackColor       =   14800597
         Caption         =   "KIRIM SMS"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin DBLiqo.jcbutton XPButton3 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Caption         =   "X"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton2 
      Height          =   375
      Left            =   720
      TabIndex        =   7
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
   Begin DBLiqo.jcbutton XPButton1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
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
   Begin VB.TextBox TxtP 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
   End
   Begin VB.ComboBox CbP 
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   6600
      Width           =   1575
   End
   Begin DBLiqo.jcbutton CmBr 
      Height          =   500
      Left            =   6840
      TabIndex        =   1
      Top             =   6240
      Width           =   1300
      _ExtentX        =   2302
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DATA BARU"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton CmCtk 
      Height          =   975
      Left            =   3960
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "CETAK REPORT"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton CmEdt 
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   6240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "EDIT DATA"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton CmHps 
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   6240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "HAPUS"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button1 
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "SMS KKI"
      UseMaskCOlor    =   -1  'True
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":2DBBE
      Height          =   5775
      Left            =   50
      TabIndex        =   18
      ToolTipText     =   "Double Click Untuk melihat Data KKI Secara Lengkap"
      Top             =   360
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   32768
      HeadLines       =   1
      RowHeight       =   27
      RowDividerStyle =   0
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
         Caption         =   "THN"
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
         EndProperty
         BeginProperty Column02 
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiCombo()
    Dim i As Integer
    With CbP
        .Clear
        For i = 0 To rsKKI.Fields.Count - 1
            .AddItem rsKKI.Fields(i).Name
        Next i
    End With
End Sub

Private Sub btAktif_Click()
    Form12.Show
End Sub

Private Sub button1_Click()
    Frame1.Visible = True
    Option1(0).Value = True
    rsIkh.Requery
    rsUst.Requery
End Sub

Private Sub button2_Click()
    Frame1.Visible = False
End Sub

Private Sub button3_Click()
Dim i As Integer
    With Form14
     .List1.Clear
     rsIkh.Filter = ""
     .Tab1.Tabs.Item(3).Selected = True
     If Option1(1).Value = True Then
        If Check1.Value = 1 Then
            rsIkh.Find "namaL='" & rsKKI!Naqib & "'"
            If Not rsIkh.EOF Then .List1.AddItem rsIkh!Hp
        End If
        If Check2.Value = 1 Then
            
            rsIkh.Find "namaL='" & rsKKI!Bendahara & "'"
            If Not rsIkh.EOF Then .List1.AddItem rsIkh!Hp
        End If
     ElseIf Option1(0).Value = True Then
        rsIkh.Filter = "NmHalaqa='" & rsKKI!namaKKI & "'"
        
        If Not rsIkh.EOF Then
            rsIkh.MoveFirst
            For i = 1 To rsIkh.RecordCount
                .List1.AddItem rsIkh!Hp
                rsIkh.MoveNext
            Next i
        End If
     
     End If
        If Check3.Value = 1 Then
            rsUst.Requery
            rsUst.Find "nama='" & rsKKI!murobbi & "'"
            If Not rsUst.EOF Then .List1.AddItem rsUst!Hp
        End If
     
        .Show
    End With
End Sub

Private Sub CbP_Click()
    If CbP.ListIndex >= 0 Then txtP.SetFocus
End Sub

Private Sub CmBr_Click()
    Baru = True
    With Form2
        .CbJen.text = "-"
        .Cbm.text = "-"
        .CbNq.text = "-"
        .CbBd.text = "-"
        
        .Show
    End With
End Sub

Private Sub CmCtk_Click()
With Form10.CRKKI
    .ReportFileName = App.Path & "\RPT\RptKKI.rpt"
    .DataFiles(0) = App.Path & "\Db.mdb"
    .WindowState = crptMaximized
    If txtP.text <> "" Then .ReplaceSelectionFormula "{DBKKI." & CbP.text & "} LIKE '*" & txtP.text & "*'"
    .Action = 1
    .Reset
End With
End Sub

Private Sub CmEdt_Click()
If Not rsKKI.EOF Then
    Baru = False
    With Form2
        .Text1(0).text = rsKKI!namaKKI
        .Text1(1).text = rsKKI!Tahun
        .Cbm.text = "Ust. " & rsKKI!murobbi
        .CbJen.text = rsKKI!Tipe
        .CbNq.text = rsKKI!Naqib
        .CbBd.text = rsKKI!Bendahara
        .cb.text = rsKKI!Hari
        .Text1(5).text = rsKKI!Waktu
         .Text1(6).text = rsKKI!Tempat
        .Text1(7).text = rsKKI!Kecamatan
        .Cbt.text = rsKKI!Jenis
        .CbJen.text = rsKKI!Tipe
        .Show
    End With
End If
End Sub

Private Sub CmEdit_Click()

End Sub

Private Sub CmHps_Click()
If Not rsKKI.EOF Then
    hapus = MsgBox("Hapus ?", vbInformation + vbYesNo, "Penghapusan")
    If hapus = vbYes Then
        rsKKI.Delete
        rsKKI.Requery
    End If
End If
End Sub

Private Sub DataGrid1_AfterUpdate()
    rsKKI.Update
End Sub

Private Sub DataGrid1_DblClick()
    Form5.Show
End Sub

Private Sub Form_Load()
    isiCombo
    rsKKI.Filter = ""
    refreshData
    Set DataGrid1.DataSource = rsKKI
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
        Check1.Value = 1
        Check2.Value = 1
        Check1.Enabled = False
        Check2.Enabled = False
    ElseIf Option1(1).Value = True Then
        Check1.Enabled = True
        Check2.Enabled = True
    End If
End Sub

Private Sub txtP_Change()
    If txtP.text = "" Then
        rsKKI.Filter = ""
        rsKKI.Requery
    Else
        rsKKI.Filter = CbP.text & " like '%" & txtP.text & "%'"
    End If
End Sub

Private Sub XPButton1_Click()
If Not rsKKI.EOF Then If Not rsKKI.AbsolutePosition <= 1 Then rsKKI.MovePrevious
End Sub

Private Sub XPbutton2_Click()
If Not rsKKI.EOF Then If Not rsKKI.AbsolutePosition >= rsKKI.RecordCount Then rsKKI.MoveNext
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub
