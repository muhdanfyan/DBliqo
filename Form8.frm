VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   Caption         =   "Data Ustadz Murobbi"
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   8550
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin DBLiqo.jcbutton Btn3 
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   7560
      Width           =   1075
      _ExtentX        =   1905
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
      Caption         =   "Hapus Data"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton Btn2 
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   7560
      Width           =   1075
      _ExtentX        =   1905
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
      Caption         =   "Edit Data"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton Btn1 
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   7560
      Width           =   1075
      _ExtentX        =   1905
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
      Caption         =   "Data Baru"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPButton4 
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   7560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      Caption         =   "Cetak Report"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtP 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   7920
      Width           =   1695
   End
   Begin VB.ComboBox CbP 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin DBLiqo.jcbutton XPButton3 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8160
      TabIndex        =   3
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
   Begin DBLiqo.jcbutton XPButton1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form8.frx":2C5B1
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Untuk melihat Data Secara Lengkap"
      Top             =   480
      Width           =   8630
      _ExtentX        =   15214
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Nama"
         Caption         =   "NAMA USTADZ"
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
         DataField       =   "Alamat"
         Caption         =   "ALAMAT LENGKAP"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
      EndProperty
   End
   Begin DBLiqo.jcbutton XPButton2 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
   Begin DBLiqo.jcbutton jcbutton1 
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      Caption         =   "SMS USTADZ"
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn1_Click()
    Baru = True
    Form7.Show
End Sub

Private Sub Btn2_Click()
    If Not rsUst.EOF Then
        Baru = False
        With Form7
            .txt(0).text = rsUst!nama
            .txt(1).text = rsUst!Alamat
            .txt(2).text = rsUst!Hp
            .Show
        End With
    End If
End Sub

Private Sub Btn3_Click()
    Dim hapus As String
    If Not rsUst.EOF Then
        hapus = MsgBox("Hapus Data ?", vbYesNo, "Hapus")
        If hapus = vbYes Then rsUst.Delete
    End If
End Sub

Private Sub CbP_Click()
    If CbP.ListIndex >= 0 Then txtP.SetFocus
End Sub

Private Sub DataGrid1_AfterUpdate()
    rsUst.Update
End Sub

Private Sub DataGrid1_DblClick()
    Form9.Show
End Sub

Private Sub Form_Load()
    Dim i As Integer
    rsUst.Filter = ""
    Set DataGrid1.DataSource = rsUst
    
    With CbP
        .Clear
        For i = 0 To rsUst.Fields.Count - 1
            .AddItem rsUst.Fields(i).Name
        Next i
    End With
    refreshData
End Sub

Private Sub jcbutton1_Click()
    Dim i As Integer
    
    Form14.List1.Clear
    rsUst.Requery
    If rsUst.EOF = False Then rsUst.MoveFirst
    For i = 1 To rsUst.RecordCount
        Form14.Tab1.Tabs(3).Selected = True
        Form14.List1.AddItem rsUst!Hp
        rsUst.MoveNext
    Next i
        Form14.Show

End Sub

Private Sub txtP_Change()
    If txtP.text = "" Then
        rsUst.Filter = ""
        rsUst.Requery
    Else
        rsUst.Filter = CbP.text & " LIKE '%" & txtP.text & "%'"
    End If
End Sub

Private Sub XPButton1_Click()
If Not rsUst.EOF Then If Not rsUst.BOF Then rsUst.MovePrevious
End Sub

Private Sub XPbutton2_Click()
If Not rsUst.EOF Then If Not rsUst.AbsolutePosition >= rsUst.RecordCount Then rsUst.MoveNext
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub

Private Sub XPButton4_Click()
    With Form10.CRUst
        .ReportFileName = App.Path & "\RPT\RptUst.rpt"
        .DataFiles(0) = App.Path & "\Db.mdb"
        .WindowState = crptMaximized
        If txtP.text <> "" Then .ReplaceSelectionFormula "{DBUst." & CbP.text & "} LIKE '*" & txtP.text & "*'"
        .Action = 1
        .Reset
    End With
End Sub
