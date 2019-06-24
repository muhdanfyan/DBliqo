VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   Caption         =   "SMS KADER"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17235
   LinkTopic       =   "Form13"
   MDIChild        =   -1  'True
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   17235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   2820
      Left            =   3840
      TabIndex        =   13
      Top             =   2120
      Visible         =   0   'False
      Width           =   4800
      Begin DBLiqo.jcbutton jcbutton8 
         Height          =   495
         Left            =   100
         TabIndex        =   15
         Top             =   2295
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         ButtonStyle     =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Balas"
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   100
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   4600
      End
      Begin DBLiqo.jcbutton jcbutton9 
         Height          =   495
         Left            =   3840
         TabIndex        =   16
         Top             =   2290
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         ButtonStyle     =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tutup"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   8400
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   7680
      Width           =   12135
      Begin DBLiqo.jcbutton jcbutton5 
         Height          =   615
         Left            =   6000
         TabIndex        =   10
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         Caption         =   "SARAN DAN PERTANYAAN"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton1 
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         BackColor       =   16761024
         Caption         =   "KIRIM SMS"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton2 
         Height          =   615
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         Caption         =   "AKTIFKAN LAYANAN"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton3 
         Height          =   615
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         Caption         =   "MATIKAN LAYANAN"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton6 
         Height          =   615
         Left            =   8040
         TabIndex        =   11
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         Caption         =   "HAPUS PESAN"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton7 
         Height          =   615
         Left            =   10080
         TabIndex        =   12
         Top             =   120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1085
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
         Caption         =   "TUTUP"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   11520
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "Semua Pesan Masuk"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Belum terbaca"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin DBLiqo.jcbutton jcbutton4 
         Height          =   615
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
         Caption         =   "BACA SMS"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   12726
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13573
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesan Keluar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesan Terkirim"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesan Masuk"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_DblClick()
Dim noHp As String
Dim namax As String
Dim text As String
    
    If TabStrip1.Tabs(3).Selected Then
        Frame3.Caption = inBox!Sendernumber
        ckNmr inBox!Sendernumber, noHp
        rsIkh.Find "hp='" & noHp & "'"
        If rsIkh.EOF Then
            rsUst.Find "Hp='" & noHp & "'"
            If rsUst.EOF Then
                namax = "Tidak dikenal"
            Else
                namax = "Ust. " & rsUst!nama
            End If
        Else
            namax = rsIkh!namaL
        End If
        
        Text1.text = inBox!Textdecoded & vbCrLf & "Pengirim : " & namax
        
        If inBox!Processed = "false" Then
            inBox!Processed = True
            inBox.Update
        End If
        
        Frame3.Visible = True
    End If
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    Dim sFilter As String
    
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        
    End Select
End Sub

Private Sub Form_Load()
    refreshData
    Set DataGrid1.DataSource = outBox
End Sub

Private Sub Form_Resize()
    TabStrip1.Width = Me.Width - 250
    DataGrid1.Width = TabStrip1.Width - 100
    Frame1.Left = TabStrip1.Width - Frame1.Width
End Sub

Private Sub jcbutton1_Click()
    With Form14
    .Show
    .Tab1.Tabs(1).Selected = True
    End With
End Sub

Private Sub jcbutton2_Click()
Dim a As String
    CreateObject("wscript.shell").run "taskkill /f /im gammu.exe"
    a = MsgBox("tampilkan Proses Layanan SMS?..", vbInformation + vbYesNo, App.EXEName)
    If a = vbYes Then
        CreateObject("wscript.shell").run "gammu.lnk"
    Else
        CreateObject("wscript.shell").run "gammu.lnk", vbHide
    End If
End Sub

Private Sub jcbutton3_Click()
    CreateObject("wscript.shell").run "taskkill /f /im gammu.exe"
    MsgBox "Matikan Layanan SMS"
End Sub

Private Sub jcbutton4_Click()
    DataGrid1_DblClick
End Sub

Private Sub jcbutton6_Click()
Dim a
a = MsgBox("Yakin ingin dihapus ? ", vbYesNo + vbInformation, "Hapus data")
If a = vbYes Then
    If TabStrip1.Tabs(1).Selected Then
        If Not outBox.EOF Then
             outBox.Delete
             outBox.Requery
        Else
            MsgBox "Data tidak ada...", vbCritical, "Maaf.."
        End If
    ElseIf TabStrip1.Tabs(2).Selected Then
        If Not sentBox.EOF Then
            sentBox.Delete
        Else
            MsgBox "Data tidak ada...", vbCritical, "Maaf.."
        End If
    ElseIf TabStrip1.Tabs(3).Selected Then
        If Not inBox.EOF Then
            inBox.Delete
        Else
            MsgBox "Data tidak ada...", vbCritical, "Maaf.."
        End If
    End If
End If
End Sub

Private Sub jcbutton7_Click()
    Unload Me
End Sub

Private Sub jcbutton8_Click()
    If Not inBox.EOF Then
        With Form14
            .Tab1.Tabs(3).Selected = True
            .List1.AddItem Frame3.Caption
            .Show
        End With
    End If
End Sub

Private Sub jcbutton9_Click()
    Frame3.Visible = False
    
    inBox.Filter = "Processed='false'"
    If inBox.EOF Then
        inBox.Filter = ""
        Option1(1).Value = True
    End If
    inBox.Requery
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value Then
        inBox.Filter = "Processed='" & False & "'"
    ElseIf Option1(1).Value Then
        inBox.Filter = ""
        inBox.Requery
    End If
    
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.Tabs(1).Selected Then
        Set DataGrid1.DataSource = outBox
        Frame1.Visible = False
        Timer1.Enabled = True
    ElseIf TabStrip1.Tabs(2).Selected Then
        Set DataGrid1.DataSource = sentBox
        Frame1.Visible = False
        Timer1.Enabled = True
    ElseIf TabStrip1.Tabs(3).Selected Then
        Set DataGrid1.DataSource = inBox
        Frame1.Visible = True
        Timer1.Enabled = False
        inBox.Filter = "Processed='" & False & "'"
        
        If inBox.EOF Then
            inBox.Filter = ""
            Option1(1).Value = True
        End If
        
        inBox.Requery
    End If
End Sub

Private Sub Timer1_Timer()
If Not outBox.EOF Then
    outBox.Requery
    sentBox.Requery
End If
'   inBox.Requery
TabStrip1.Tabs(1).Caption = "Pesan Keluar[" & outBox.RecordCount & "]"
TabStrip1.Tabs(2).Caption = "Pesan Terkirim[" & sentBox.RecordCount & "]"
TabStrip1.Tabs(3).Caption = "Pesan Masuk[" & inBox.RecordCount & "]"

End Sub
