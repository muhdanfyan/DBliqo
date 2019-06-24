VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form12 
   BorderStyle     =   0  'None
   Caption         =   "DAFTAR KTI BINAAN WAHDAH ISLAMIYAH MAKASSAR"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13740
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   11040
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   12720
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form12.frx":3493B
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   12912
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   22
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "NamaKKI"
         Caption         =   "Nama KTI"
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
         DataField       =   "Murobbi"
         Caption         =   "Murobbi"
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
         DataField       =   "Naqib"
         Caption         =   "Naqib"
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
         DataField       =   "aktif"
         Caption         =   "Aktif"
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
         DataField       =   "Taktif"
         Caption         =   "Nonaktif"
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
      BeginProperty Column06 
         DataField       =   "Hari"
         Caption         =   "Hari"
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
         DataField       =   "Waktu"
         Caption         =   "Waktu"
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
         DataField       =   "Jenis"
         Caption         =   "Jenis"
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
         DataField       =   "Tahun"
         Caption         =   "Tahun"
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
         DataField       =   "Tipe"
         Caption         =   "Tipe"
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
         DataField       =   "wkt"
         Caption         =   "wkt"
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
         DataField       =   "Ket"
         Caption         =   "Ket"
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
         EndProperty
         BeginProperty Column03 
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
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrUpKKI 
      Left            =   4560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin DBLiqo.jcbutton BtTutup 
      Height          =   615
      Left            =   12390
      TabIndex        =   5
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "TUTUP"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   7920
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin DBLiqo.jcbutton button1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "<<"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button2 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   ">>"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton BtnCetak 
      Height          =   615
      Left            =   11040
      TabIndex        =   6
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   "CETAK"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub valDat()
    rsAkt.Filter = "wkt='" & Text1.text & "-" & Combo1.ListIndex + 1 & "' AND aktif>0"
End Sub

Sub isiCombo()
Dim i As Integer
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
    
    
    Combo2.Clear
For i = 0 To rsAkt.Fields.Count - 1
    Combo2.AddItem rsAkt.Fields(i).Name
Next i
End Sub

Private Sub BtnCetak_Click()
With CrUpKKI
    .ReportFileName = App.Path & "\RPT\RptupKKI.rpt"
    .DataFiles(0) = App.Path & "\Db.mdb"
    .WindowState = crptMaximized
    If Text1.text <> "" Then .ReplaceSelectionFormula "{UpKKI.wkt} = '" & Text1.text & "-" & Combo1.ListIndex + 1 & "'"
    .Action = 1
    .Reset
End With
End Sub

Private Sub BtTutup_Click()
    Unload Me
End Sub

Private Sub button1_Click()
If Not rsAkt.EOF Then If Not rsAkt.AbsolutePosition <= 1 Then rsAkt.MovePrevious
End Sub

Private Sub button2_Click()
If Not rsAkt.EOF Then If Not rsAkt.AbsolutePosition >= rsAkt.RecordCount Then rsAkt.MoveNext
End Sub

Private Sub Combo1_Change()
    valDat
End Sub

Private Sub Combo1_Click()
    valDat
End Sub

Private Sub Combo2_Click()
    If Combo2.ListIndex > -1 Then Text2.SetFocus
End Sub

Private Sub DataGrid1_AfterUpdate()
    rsupIk.Update
End Sub

Private Sub Form_Load()
Dim i As Integer
    isiCombo
    Combo1.ListIndex = Format(Date, "m") - 1
    Text1.text = Format(Date, "YYYY")
    valDat
    rsKKI.Filter = ""
    rsKKI.MoveFirst
    For i = 1 To rsKKI.RecordCount
        rsupKKI.Filter = "NamaKKI='" & rsKKI!namaKKI & "' AND wkt='" & Format(Date, "YYYY") & "-" & Format(Date, "m") & "'"
        
        If rsupKKI.EOF Then rsupKKI.AddNew
        
        rsupKKI!namaKKI = rsKKI!namaKKI
        aktif.Filter = "KKI='" & rsKKI!namaKKI & "' AND wkt='" & Format(Date, "YYYY") & "-" & Format(Date, "m") & "' AND keaktifan='Aktif'"
        rsupKKI!aktif = aktif.RecordCount
        Taktif.Filter = "KKI='" & rsKKI!namaKKI & "' AND wkt='" & Format(Date, "YYYY") & "-" & Format(Date, "m") & "' AND keaktifan='Tidak Aktif' OR keaktifan='Kurang'"
        rsupKKI!Taktif = Taktif.RecordCount
        rsupKKI!wkt = Format(Date, "YYYY") & "-" & Format(Date, "m")
        rsupKKI.Update
        rsupKKI.Requery
        rsKKI.MoveNext
    Next i
    rsAkt.Requery
    Set DataGrid1.DataSource = rsAkt
    refreshData
End Sub

Private Sub Text1_Change()
    valDat
End Sub

Private Sub Text2_Change()
    If Text2.text = "" Then
        valDat
        rsAkt.Requery
    Else
        rsAkt.Filter = rsAkt.Filter & " AND " & Combo2.text & " like '%" & Text2.text & "%'"
        
    End If
End Sub
