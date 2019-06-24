VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Data Ikhwah"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   13335
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton jcbutton3 
      Height          =   345
      Left            =   480
      TabIndex        =   12
      Top             =   8520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   609
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
      Caption         =   "Filter secara lengkap"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton jcbutton1 
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   7800
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
      Caption         =   "SMS Ikhwah"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox TxtP 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   8160
      Width           =   1935
   End
   Begin VB.ComboBox CBp 
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   8160
      Width           =   1935
   End
   Begin DBLiqo.jcbutton CmHps 
      Height          =   495
      Left            =   12360
      TabIndex        =   7
      Top             =   7920
      Width           =   900
      _ExtentX        =   1588
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
      Caption         =   "HAPUS"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton CmBr 
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   7920
      Width           =   900
      _ExtentX        =   1588
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
      Caption         =   "DATA BARU"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton3 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   12840
      TabIndex        =   4
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
      Caption         =   "x"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton4 
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      Caption         =   "CETAK REPORT"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton1 
      Height          =   375
      Left            =   40
      TabIndex        =   1
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":3D1F3
      Height          =   7500
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Double Click Untuk melihat Data Ikhwah Secara Lengkap"
      Top             =   360
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13229
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   13977088
      HeadLines       =   1
      RowHeight       =   20
      RowDividerStyle =   6
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
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "NamaL"
         Caption         =   "NAMA LENGKAP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NamaKunn"
         Caption         =   "KUNNIYAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Tempat"
         Caption         =   "Tempat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AlamatD"
         Caption         =   "AlamatD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "AlamatM"
         Caption         =   "ALAMAT (MAKASSAR)"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Pendidikan"
         Caption         =   "Pendidikan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "NamaS"
         Caption         =   "PENDIDIKAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Angkt"
         Caption         =   "Angkt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Fak"
         Caption         =   "Fak"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Jur"
         Caption         =   "Jur"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Bkt"
         Caption         =   "Bkt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "USaud"
         Caption         =   "USaud"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "JSaud"
         Caption         =   "JSaud"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "RSD"
         Caption         =   "RSD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "RSMP"
         Caption         =   "RSMP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "RSMA"
         Caption         =   "RSMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "POrg"
         Caption         =   "POrg"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "NmHalaqa"
         Caption         =   "NAMA HALAQAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Tingkat"
         Caption         =   "TINGKAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "Ayah"
         Caption         =   "Ayah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Ibu"
         Caption         =   "Ibu"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "hp"
         Caption         =   "NO TELP/HP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Gol"
         Caption         =   "Gol"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin DBLiqo.jcbutton XPButton2 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   720
      _ExtentX        =   1270
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
   Begin DBLiqo.jcbutton CmEdt 
      Height          =   495
      Left            =   11400
      TabIndex        =   6
      Top             =   7920
      Width           =   900
      _ExtentX        =   1588
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
      Caption         =   "EDIT"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton jcbutton2 
      Height          =   495
      Left            =   9000
      TabIndex        =   11
      Top             =   8280
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
      Caption         =   "SMS dalam Grid"
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiCombo()
Dim i As Integer
    With CbP
        .Clear
        For i = 0 To rsIkh.Fields.Count - 1
            .AddItem rsIkh.Fields(i).Name
        Next i
    End With
End Sub

Private Sub CbP_Click()
    If CbP.ListIndex >= 0 Then txtP.SetFocus
End Sub

Private Sub CmBr_Click()
    Baru = True
    Form4.Show
End Sub

Private Sub CmEdt_Click()
If Not rsIkh.EOF Then
    Baru = False
    With Form4
        .txt(0).text = rsIkh!namaL
        .txt(1).text = rsIkh!NamaKunn
        .txt(2).text = rsIkh!Tempat
        If rsIkh!tanggal <> Empty Then .DT1.Value = rsIkh!tanggal
        .txt(3).text = rsIkh!AlamatD
        .txt(4).text = rsIkh!AlamatM
        .CbPen.text = rsIkh!Pendidikan
        .txt(17).text = rsIkh!NamaS
        .txt(18).text = rsIkh!jur
        .txt(5).text = rsIkh!Fak
        .txt(6).text = rsIkh!Angkt
        .txt(7).text = rsIkh!Usaud
        .txt(8).text = rsIkh!Jsaud
        .txt(9).text = rsIkh!RSD
        .txt(10).text = rsIkh!RSMP
        .txt(11).text = rsIkh!RSMA
        .txt(12).text = rsIkh!Bkt
        .txt(13).text = rsIkh!Porg
        .txt(14).text = rsIkh!Ayah
        .txt(15).text = rsIkh!Ibu
        .txt(16).text = rsIkh!Hp
        .CbH1.text = rsIkh!NmHalaqa
        .lblT.Caption = rsIkh!Tingkat
        .Show
    End With
End If
End Sub

Private Sub CmHps_Click()

If Not rsIkh.EOF Then
    hapus = MsgBox("Hapus ?", vbInformation + vbYesNo, "Penghapusan")
    If hapus = vbYes Then
        rsIkh.Delete
        rsIkh.Requery
    End If
End If
End Sub

Private Sub DataGrid1_AfterUpdate()
    rsIkh.Update
End Sub

Private Sub DataGrid1_DblClick()
    Form6.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    rsKKI.Filter = ""
    refreshData
End Sub

Private Sub Form_Load()
    isiCombo
    rsIkh.Filter = ""
    Set DataGrid1.DataSource = rsIkh
    refreshData
End Sub

Private Sub Label1_Click()
    XPFrame1_Click
End Sub

Private Sub jcbutton3_Click()
    With Form4
        .XPButton1.Caption = "TAMPIILKAN FILTER"
        .Show
    End With
End Sub

Private Sub jcbutton1_Click()
    If Not rsIkh.EOF Then
        With Form14
            .Tab1.Tabs(3).Selected = True
            .List1.AddItem rsIkh!Hp
            .Show
        End With
    End If
End Sub

Private Sub jcbutton2_Click()
    Dim i As Integer
    Form14.List1.Clear
    rsIkh.Requery
    If rsIkh.EOF = False Then rsIkh.MoveFirst
    For i = 1 To rsIkh.RecordCount
        Form14.Tab1.Tabs(3).Selected = True
        Form14.List1.AddItem rsIkh!Hp
        rsIkh.MoveNext
    Next i
        Form14.Show
End Sub

Private Sub txtP_Change()
    If txtP.text = "" Then
        rsIkh.Filter = ""
        rsIkh.Requery
    Else
        rsIkh.Filter = CbP.text & " Like '%" & txtP.text & "%'"
    End If
End Sub

Private Sub XPButton1_Click()
    If Not rsIkh.EOF Then If Not rsIkh.AbsolutePosition <= 1 Then rsIkh.MovePrevious
End Sub

Private Sub XPbutton2_Click()
    If Not rsIkh.EOF Then If Not rsIkh.AbsolutePosition >= rsIkh.RecordCount Then rsIkh.MoveNext
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub

Private Sub XPFrame1_Click()
    XPFrame1.BackStyle = Transparent
End Sub

Private Sub XPButton4_Click()
    With Form10.CRIkh
        .ReportFileName = App.Path & "\RPT\RptIkh.rpt"
        .DataFiles(0) = App.Path & "\Db.mdb"
        .WindowState = crptMaximized
        If txtP.text <> "" Then .ReplaceSelectionFormula "{DBIkhwah." & CbP.text & "} LIKE '*" & txtP.text & "*'"
        .Action = 1
        .Reset
    End With
End Sub
