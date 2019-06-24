VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Cetak Report"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   5415
      Begin VB.Frame Frm 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4815
         Begin Crystal.CrystalReport CRUst 
            Left            =   4440
            Top             =   1560
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Semua"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Kecamatan"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   2
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Marhalah"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   4
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            Top             =   960
            Width           =   2175
         End
         Begin VB.ComboBox CbTk 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   1320
            Width           =   2175
         End
      End
      Begin VB.ComboBox CbDt 
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
   End
   Begin DBLiqo.jcbutton XPButton3 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
      Caption         =   "X"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton1 
      Height          =   3015
      Left            =   5640
      TabIndex        =   6
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   5318
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
   Begin Crystal.CrystalReport CRIkh 
      Left            =   900
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CRKKI 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbDt_Click()
    If CbDt.ListIndex = 0 Then
        Frm.Visible = True
        Opt(0).Value = True
    Else
        Frm.Visible = False
    End If
End Sub

Private Sub Form_Load()
    With CbDt
        .Clear
        .AddItem "Data KKI"
        .AddItem "Data Ikhwah"
        .AddItem "Data Ustadz"
    End With
    With CbTk
        .Clear
        .AddItem "PRA TA'RIF"
        .AddItem "TA'RIFIYAH"
        .AddItem "TAKWINIYAH"
    End With
    refreshData
End Sub

Private Sub Opt_Click(Index As Integer)
    If Opt(1).Value = True Then
        txt.Visible = True
        txt.SetFocus
    Else
        txt.Visible = False
    End If
    
    If Opt(2).Value = True Then
        CbTk.Visible = True
        CbTk.SetFocus
    Else
        CbTk.Visible = False
    End If
End Sub

Private Sub XPButton1_Click()
    If CbDt.ListIndex = 0 Then
        With CRKKI
            .ReportFileName = App.Path & "\RPT\RptKKI.rpt"
            .DataFiles(0) = App.Path & "\Db.mdb"
            .WindowState = crptMaximized
            If Opt(1).Value = True Then
                .ReplaceSelectionFormula "{DBKKI.Kecamatan} like '*" & txt.text & "*'"
            ElseIf Opt(2).Value = True Then
                If CbTk.ListIndex = 0 Then
                    .ReplaceSelectionFormula "{DBKKI.Jenis} like 'PRA*'"
                ElseIf CbTk.ListIndex = 1 Then
                    .ReplaceSelectionFormula "{DBKKI.Jenis} like '*RIFIYAH'"
                ElseIf CbTk.ListIndex = 2 Then
                    .ReplaceSelectionFormula "{DBKKI.Jenis} like '*WINIYAH'"
                End If
            End If
            .Action = 1
            .Reset
        End With
      
    ElseIf CbDt.ListIndex = 1 Then
        With CRIkh
            .ReportFileName = App.Path & "\RPT\RptIkh.rpt"
            .DataFiles(0) = App.Path & "\Db.mdb"
            .WindowState = crptMaximized
            .Action = 1
            .Reset
        End With
    ElseIf CbDt.ListIndex = 2 Then
        With CRUst
            .ReportFileName = App.Path & "\RPT\RptUst.rpt"
            .DataFiles(0) = App.Path & "\Db.mdb"
            .WindowState = crptMaximized
            .Action = 1
            .Reset
        End With
    End If
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub

