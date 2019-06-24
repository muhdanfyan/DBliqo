VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormUtama 
   Caption         =   "Aplikasi DataBase Liqo Tarbiyah"
   ClientHeight    =   9180
   ClientLeft      =   135
   ClientTop       =   -465
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "FormUtama.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   9765
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   8040
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7858
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
               LCID            =   1057
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
               LCID            =   1057
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Konfirmasi Kehadiran"
      Height          =   2655
      Left            =   3600
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   4455
      Begin DBLiqo.jcbutton jcbutton4 
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   2160
         Width           =   4455
         _ExtentX        =   7858
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
         Caption         =   "Tutup"
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Label5"
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   3735
      End
   End
   Begin DBLiqo.jcbutton jcbutton1 
      Height          =   975
      Left            =   10920
      TabIndex        =   9
      Top             =   240
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1720
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
      Caption         =   ""
      Picture         =   "FormUtama.frx":8D685
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   1800
      Top             =   5040
   End
   Begin DBLiqo.jcbutton jcbutton2 
      Height          =   495
      Left            =   16800
      TabIndex        =   8
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pesan Masuk"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button1 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1720
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
      Caption         =   "Data Halaqoh"
      Picture         =   "FormUtama.frx":9124D
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton XPbutton1 
      Cancel          =   -1  'True
      Height          =   1095
      Left            =   12840
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
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
      Caption         =   ""
      Picture         =   "FormUtama.frx":952C9
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   5040
   End
   Begin DBLiqo.jcbutton button2 
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1720
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
      Caption         =   "Data Ikhwah"
      Picture         =   "FormUtama.frx":99C0E
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button3 
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1720
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
      Caption         =   "Data Ustadz"
      Picture         =   "FormUtama.frx":9D79E
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button4 
      Height          =   975
      Left            =   6600
      TabIndex        =   5
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1720
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
      Caption         =   "Cetak Report"
      Picture         =   "FormUtama.frx":A1289
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button5 
      Height          =   975
      Left            =   8760
      TabIndex        =   6
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1720
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
      Caption         =   "SMS Gateway"
      Picture         =   "FormUtama.frx":A4D5C
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton jcbutton3 
      Height          =   495
      Left            =   16800
      TabIndex        =   10
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Request Halaqah"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   17160
      TabIndex        =   11
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   17160
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   8400
      Width           =   720
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rn As Currency
Dim hpF As String
Public xrskki As New ADODB.Recordset
Public xrsust As New ADODB.Recordset
Public xrsIkh As New ADODB.Recordset
Public xDaft As New ADODB.Recordset

Sub PesanMasuk()
    iB.Requery
    iB.Filter = "Processed='" & False & "'"
    If iB.EOF Then
        Label2.Caption = "Tidak ada pesan masuk"
    Else
        Label2.Caption = ""
        Label2.Caption = "[Pesan Yang Baru Masuk] >> "
        iB.MoveFirst
        For it = 1 To iB.RecordCount
            Label2.Caption = Label2.Caption & "-" & iB!Sendernumber & " : " & iB!Textdecoded & vbCrLf
            iB.MoveNext
        Next it
    End If
    
    With xDaft
        .Requery
        Label3.Caption = ""
        .Filter = "NmHalaqa='-'"
      If xDaft.EOF Then
        Label3.Caption = "Tidak Ada Pendaftar.."
      Else
        For it = 1 To .RecordCount
            Label3.Caption = Label3.Caption & "- " & !namaL & "(Hp: " & !Hp & ")" & vbCrLf
        Next it
      End If
    End With
    
End Sub

Sub smsPart(txSMS As String, nhp As String)
Dim i, j, k, l, n, m As Integer
    n = 1
    m = 0
    k = Len(txSMS) / 152 + 1
    For j = 1 To k
        n = n + 152
        If j = k Then
            l = k Mod 152
        Else
            l = 152
        End If
            kirimSMS "info" & i & "/" & k & ":" & Mid(txSMS, n, l), nhp
    Next j
End Sub

Public Sub cekupMur(t As String)
    With xinBox
        .Filter = "TextDecoded like '" & t & "(%' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
       ' Timer2.Enabled = False
        j = Len(t & "(")
        Do Until (d = ")")
            d = Mid(!Textdecoded, j + 1, 1)
            j = j + 1
        Loop
        
        Krit = Mid(!Textdecoded, Len(t & "(") + 1, j - Len(t & "(") - 1)
        
        xrskki.Filter = "NamaKKI='" & Krit & "'"
        
        'MsgBox "[" & Krit & "]"
        
        If Not xrskki.EOF Then
            xrsust.Filter = "Nama='" & xrskki!murobbi & "'"
            
            ckNmr !Sendernumber, hpF
            
            If hpF = xrsust!Hp Then
                If t = "gtHr" Then
                    xrskki!Hari = Right(!Textdecoded, Len(!Textdecoded) - j)
                ElseIf t = "gtWkt" Then
                    xrskki!Waktu = Right(!Textdecoded, Len(!Textdecoded) - j)
                ElseIf t = "gtT4" Then
                    xrskki!Tempat = Right(!Textdecoded, Len(!Textdecoded) - j)
                End If
                xrskki.Update
                sms "Data Halaqoh " & Krit & " telah berhasil terupdate.. ", !Sendernumber
            End If
        Else
            sms "Data Halaqah ini tidak ada..", !Sendernumber
        End If
            !Processed = True
            .MoveNext
            rsKKI.Requery
    End If
End With

End Sub

Sub cekKonfirm()

Dim i As Integer
Dim a1, a2, a3 As Integer

    'sms konfirmasi
With xinBox
    
    .Filter = "TextDecoded like 'ceck(%' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        j = Len("konf(")
        Do Until (d = ")")
            d = Mid(!Textdecoded, j + 1, 1)
            j = j + 1
        Loop
        
        Krit = Mid(!Textdecoded, Len("konf(") + 1, j - Len("konf(") - 1)
        xrskki.Filter = "NamaKKI='" & Krit & "'"
        

        If Not xrskki.EOF Then
            xrsust.Filter = "Nama='" & xrskki!murobbi & "'"
            xrsIkh.Filter = "NmHalaqa='" & xrskki!namaKKI & "'"
            
            If Mid(!Sendernumber, 1, 3) = "+62" Then hpF = "0" & Right(!Sendernumber, Len(!Sendernumber) - Len("+62"))
            
            If hpF = xrsust!Hp Then
                rsKonf.Filter = "dari='mrb" & Krit & "' AND konf='y'"
                a1 = rsKonf.RecordCount
                rsKonf.Filter = "dari='mrb" & Krit & "' AND konf='t'"
                a2 = rsKonf.RecordCount
                rsKonf.Filter = "dari='mrb" & Krit & "' AND konf='-'"
                a3 = rsKonf.RecordCount
                rsKonf.Filter = ""
                sms "ikhwh yg dtang : " & a1 & " orang; yg tdk dtg " & a2 & " orang; blum konfirmasi : " & a3 & " orang", !Sendernumber
            End If
        Else
            sms "Data Halaqah ini tidak ada..", !Sendernumber
        End If
            !Processed = True
            .MoveNext
    End If
End With
End Sub

Sub Konfirmasi()

Dim i As Integer
    'sms konfirmasi
With xinBox
    
    .Filter = "TextDecoded like 'konf(%' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        j = Len("konf(")
        Do Until (d = ")")
            d = Mid(!Textdecoded, j + 1, 1)
            j = j + 1
        Loop
        
        Krit = Mid(!Textdecoded, Len("konf(") + 1, j - Len("konf(") - 1)
        xrskki.Filter = "NamaKKI='" & Krit & "'"
        
        'MsgBox "[" & Krit & "]"
        
        rsKonf.Filter = "kki='" & Krit & "'"
        
        For i = 1 To rsKonf.RecordCount
            rsKonf.Delete
            rsKonf.Requery
        Next i
        
        If Not xrskki.EOF Then
            xrsust.Filter = "Nama='" & xrskki!murobbi & "'"
            xrsIkh.Filter = "NmHalaqa='" & xrskki!namaKKI & "'"
            
            If Mid(!Sendernumber, 1, 3) = "+62" Then hpF = "0" & Right(!Sendernumber, Len(!Sendernumber) - Len("+62"))
            
            If hpF = xrsust!Hp Then
                For k = 1 To xrsIkh.RecordCount
                    sms "konfirm[spasi] <y/t> " & Krit & ": " & Right(!Textdecoded, Len(!Textdecoded) - j), xrsIkh!Hp
                    inputKonf Right(!Textdecoded, Len(!Textdecoded) - j), Now, xrsIkh!Hp, xrskki!namaKKI, "-", "mrb" & Krit
                    xrsIkh.MoveNext
                Next k
                sms "Permintaan diterima.. ikhwah " & Krit & " akan di hubungi..", !Sendernumber
            End If
        Else
            sms "Data Halaqah ini tidak ada..", !Sendernumber
        End If
            !Processed = True
            .MoveNext
    End If
End With
End Sub

Sub smKon(c As String)
With xinBox
    .Filter = "TextDecoded='" & c & "' AND Processed=false"
    
    If Not .EOF Then
        .MoveFirst
        Krit = !Textdecoded
        ckNmr !Sendernumber, hpF
        
        rsKonf.Filter = "dari='dpc' AND hp='" & hpF & "'"
        
        If Not rsKonf.EOF Then
            rsKonf!konf = Krit
            rsKonf.Update
        End If
        !Processed = True
'        .MoveNext
    End If
End With
End Sub


Private Sub DatHal_Click()
    Form5.Show
End Sub

Private Sub DatIkh_Click()
    Form6.Show
End Sub

Private Sub button1_Click()
Form1.Show
End Sub

Private Sub button2_Click()
Form3.Show
End Sub

Private Sub button3_Click()
Form8.Show
End Sub

Private Sub button4_Click()
    Form10.Show
End Sub

Private Sub exit_Click()
    tanya = MsgBox("Yakin Ingin Keluar ? ", vbYesNo + vbInformation, "Keluar dari Program")
    If tanya = vbYes Then End
End Sub

Private Sub button5_Click()
    Form13.Show
End Sub

Private Sub DataGrid1_DblClick()
    Frame2.Visible = False
End Sub

Private Sub Form_Load()
    
    Dim it As Integer
    If Not xrskki.State = 1 Then xrskki.Open "Select * from dbKKI ORDER BY NamaKKI ASC", Konek, adOpenDynamic, adLockOptimistic
    If Not xrsust.State = 1 Then xrsust.Open "Select * from dbUst ORDER BY Nama ASC", Konek, adOpenDynamic, adLockOptimistic
    If Not xrsIkh.State = 1 Then xrsIkh.Open "Select * from dbIkhwah ORDER BY NamaL ASC", Konek, adOpenDynamic, adLockOptimistic
    If Not xDaft.State = 1 Then xDaft.Open "Select * from dbIkhwah ORDER BY NamaL ASC", Konek, adOpenDynamic, adLockOptimistic
    
End Sub

Private Sub InputHalaqah_Click()
    Form1.Show
End Sub

Private Sub InputIkhwah_Click()
    Form3.Show
End Sub

Private Sub Form_Resize()
    XPButton1.Left = Me.Width - 1500
    XPButton1.Top = Me.Height - 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tanya = MsgBox("Yakin Ingin Keluar ? ", vbYesNo + vbInformation, "Keluar dari Program")
    If tanya = vbYes Then
        On Error Resume Next
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Frame2_Click()
    Frame2.Visible = False
End Sub

Private Sub jcbutton1_Click()
    MDI.Hide
    Form16.Width = 4005
    Form16.Height = 7470
    Form16.Top = Screen.Height - Form16.Height - 200
    Form16.Left = Screen.Width - Form16.Width - 50
    Form16.Show
End Sub

Private Sub jcbutton2_Click()
    With Form13
    .TabStrip1.Tabs(3).Selected = True
    .Option1(0).Value = True
    End With
End Sub

Private Sub jcbutton3_Click()
    With Form3
        .CbP.text = "NmHalaqa"
        .txtP.text = "-"
        .Show
    End With
    
End Sub

Private Sub jcbutton4_Click()
    Frame1.Visible = False
End Sub

Private Sub Label4_Click()
    Frame2.Visible = True
    rsKonf.Filter = "dari='dpc' AND konf='y'"
    Set DataGrid1.DataSource = rsKonf
End Sub

Private Sub Label5_Click()
    Frame2.Visible = True
    rsKonf.Filter = "dari='dpc' AND konf='t'"
    Set DataGrid1.DataSource = rsKonf
End Sub

Private Sub Label6_Click()
    Frame2.Visible = True
    rsKonf.Filter = "dari='dpc' AND konf='-'"
    Set DataGrid1.DataSource = rsKonf
End Sub

Private Sub Timer1_Timer()
    Dim i  As Integer
    Label1.Caption = "Waktu    : " & Time & vbCrLf & "Tanggal : " & Format(Date, "dd-MM-YYYY")
    PesanMasuk
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "DHalaq" Then
        Form1.Show
    ElseIf Button.Key = "DIkh" Then
        Form3.Show
    ElseIf Button.Key = "DUst" Then
        Form8.Show
    ElseIf Button.Key = "Rpt" Then
        Form10.Show
    End If
End Sub

Private Sub Timer2_Timer()

Dim nmKKI, Krit, d As String
Dim i, j, k As Integer
Dim tSMS As String
With xinBox
    .Filter = "processed=false"
    
    If .EOF Then .Requery
    
    .Filter = "TextDecoded like 'jdw %' AND processed=false"
    If Not .EOF Then
        .MoveFirst
        'Timer2.Enabled = False
        For i = 1 To .RecordCount
            nmKKI = Right(!Textdecoded, Len(!Textdecoded) - 4)
            xrskki.Filter = "NamaKKI='" & nmKKI & "'"
            If xrskki.EOF Then
                kirimSMS "Data Halaqoh yg anda input tidak ada...", !Sendernumber
            Else
                tSMS = "KKI : " & nmKKI & "; wkt : " & xrskki!Hari & ", " & xrskki!Waktu & "; t4: " & xrskki!Tempat
                If Len(tSMS) <= 158 Then
                    kirimSMS tSMS, !Sendernumber
                Else
                    smsPart tSMS, !Sendernumber
                End If
            End If
            !Processed = True
            .MoveNext
        Next i
    End If

'mendaftar tarbiyah
    .Filter = "TextDecoded like 'daftar %' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        Krit = Right(!Textdecoded, Len(!Textdecoded) - Len("daftar "))
        ckNmr !Sendernumber, hpF
        
        xrsIkh.Filter = "Hp='" & hpF & "'"
        
        If xrsIkh.EOF Then
            With xrsIkh
            .AddNew
            !namaL = Krit
            !NamaKunn = "-"
            !Tempat = "-"
            !tanggal = Date
            !AlamatD = "-"
            !AlamatM = "-"
            !Pendidikan = "-"
            !NamaS = "-"
            !Fak = "-"
            !jur = "-"
            !Angkt = "-"
            !Bkt = "-"
            !Usaud = "-"
            !Jsaud = "-"
            !RSD = "-"
            !RSMP = "-"
            !RSMA = "-"
            !Porg = "-"
            !Tingkat = "-"
            !NmHalaqa = "-"
            !Ayah = "-"
            !Ibu = "-"
            !Hp = hpF
            !Gol = "-"
            .Update
            .Requery
            xrskki.Requery
            End With
            sms "Selamat.. anda telah terdaftar dalam tarbiyah... ketik halkec[spasi] <nama_kecamatan> utk melihat halaqah yang terdaftar di kecmtan tersebut", hpF
        End If
            !Processed = True
            .MoveNext
    End If

'cek identitas murobbi
    .Filter = "TextDecoded like 'MUR %' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        'Timer2.Enabled = False
        For i = 1 To .RecordCount
            nmKKI = Right(!Textdecoded, Len(!Textdecoded) - 4)
            xrskki.Filter = "NamaKKI='" & nmKKI & "'"
            If xrskki.EOF Then
                kirimSMS "Data Halaqoh yg anda input tidak ada...", !Sendernumber
            Else
                xrsust.Filter = "Nama='" & xrskki!murobbi & "'"
                tSMS = "KKI : " & nmKKI & "; Murabbi: Ust. " & xrskki!murobbi & "; Hp:" & xrsust!Hp & "; Alamat: " & xrsust!Alamat
                
                If Len(tSMS) <= 158 Then
                    kirimSMS tSMS, !Sendernumber
                Else
                    smsPart tSMS, !Sendernumber
                End If
            End If
            
            !Processed = True
            .MoveNext
        Next i
    End If
    
'konfirmasi binaan

    .Filter = "TextDecoded like 'konfirm %' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        Krit = Right(!Textdecoded, Len(!Textdecoded) - 8)
        ckNmr !Sendernumber, hpF
        xrsIkh.Find "hp='" & hpF & "'"
        If Not xrsIkh.EOF Then
            rsKonf.Filter = "kki='" & xrsIkh!NmHalaqa & " ' AND hp='" & hpF & "'"
            If Not rsKonf.EOF Then
                rsKonf!konf = Krit
                rsKonf.Update
            End If
        End If
        !Processed = True
        .MoveNext
        rsKonf.Filter = ""
        
    End If

    .Filter = "TextDecoded like 'halkec %' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        
        For i = 1 To .RecordCount
            Krit = Right(!Textdecoded, Len(!Textdecoded) - Len("halkec "))
            xrskki.Filter = "Kecamatan='" & Krit & "' AND Tipe='Umum'"
            If xrskki.EOF Then
                kirimSMS "Data Halaqoh Wilayah untuk daerah " & Krit & " tidak ada...", !Sendernumber
            Else
                tSMS = "KTI Wil " & Krit & ":"
                xrskki.MoveFirst
                For j = 1 To xrskki.RecordCount
                    tSMS = tSMS & rsKKI!namaKKI & "; "
                Next j
                tSMS = tSMS & "ketik jdw[spasi]<namaKKI> unk lihat jadwal"
                If Len(tSMS) <= 158 Then
                    kirimSMS tSMS, !Sendernumber
                Else
                    smsPart tSMS, !Sendernumber
                End If
            End If
            !Processed = True
            .MoveNext
        Next i
       ' Timer2.Enabled = True
    End If
    
    .Filter = "TextDecoded like 'halkmps %' AND processed=false"
    If Not .EOF Then
        .MoveFirst
       ' Timer2.Enabled = False
        For i = 1 To .RecordCount
            Krit = Right(!Textdecoded, Len(!Textdecoded) - Len("halkmps "))
            xrskki.Filter = "Kecamatan='" & Krit & "' AND Tipe='Kampus'"
            
            If xrskki.EOF Then
                kirimSMS "Data KTI untuk Kampus " & Krit & " tidak ada...", !Sendernumber
            Else
                tSMS = "KTI Wil " & Krit & ":"
                xrskki.MoveFirst
                For j = 1 To xrskki.RecordCount
                    tSMS = tSMS & xrskki!namaKKI & "; "
                    xrskki.MoveNext
                Next j
                sms tSMS, !Sendernumber
            End If
            !Processed = True
            .MoveNext
        Next i
    End If
'sms murobbi ke ikhwah
    .Filter = "TextDecoded like 'sms(%' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
       ' Timer2.Enabled = False
        j = Len("sms(")
        Do Until (d = ")")
            d = Mid(!Textdecoded, j + 1, 1)
            j = j + 1
        Loop
        
        Krit = Mid(!Textdecoded, Len("sms(") + 1, j - Len("sms(") - 1)
        xrskki.Filter = "NamaKKI='" & Krit & "'"
        
        'MsgBox "[" & Krit & "]"
        
        If Not xrskki.EOF Then
            xrsust.Filter = "Nama='" & xrskki!murobbi & "'"
            xrsIkh.Filter = "NmHalaqa='" & xrskki!namaKKI & "'"
            
            If Mid(!Sendernumber, 1, 3) = "+62" Then hpF = "0" & Right(!Sendernumber, Len(!Sendernumber) - Len("+62"))
            
            If hpF = xrsust!Hp Then
                For k = 1 To xrsIkh.RecordCount
                    sms "Mrb " & Krit & ": " & Right(!Textdecoded, Len(!Textdecoded) - j), xrsIkh!Hp
                    xrsIkh.MoveNext
                Next k
                
                sms "Permintaan diterima.. ikhwah " & Krit & " akan di hubungi..", !Sendernumber
            End If
        Else
            sms "Data Halaqah ini tidak ada..", !Sendernumber
        End If
            !Processed = True
            .MoveNext
    End If
    
    For i = 0 To xrsIkh.Fields.Count - 1
        .Filter = "textDecoded like 'Up" & xrsIkh.Fields(i).Name & " %' AND Processed=False"
        If Not .EOF Then
            xrsIkh.Filter = ""
            xrsIkh.Requery
            ckNmr !Sendernumber, hpF
            xrsIkh.Find "Hp='" & hpF & "'"
            
            If Not xrsIkh.EOF Then
                xrsIkh.Fields(i).Value = Right(!Textdecoded, Len(!Textdecoded) - Len("up" & xrsIkh.Fields(i).Name & " "))
                xrsIkh.Update
                xrsIkh.Requery
            End If
            
            !Processed = True
        End If
    Next i
    
    Konfirmasi
    
    cekKonfirm
    
    cekupMur ("gtHr")
    
    cekupMur ("gtwkt")
    
    cekupMur ("gtT4")

'konfirmasi umum
    smKon "y"
    smKon "t"
    
End With

End Sub

Private Sub XPButton1_Click()
    tanya = MsgBox("Yakin Ingin Keluar ? ", vbYesNo + vbInformation, "Keluar dari Program")
    If tanya = vbYes Then End
End Sub

Private Sub XPbutton2_Click()
    Form11.Show
End Sub
