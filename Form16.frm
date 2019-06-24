VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   0  'None
   Caption         =   "Database Liqo"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   6240
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   0
      Top             =   6240
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   3855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   3915
   End
   Begin DBLiqo.jcbutton jcbutton2 
      Height          =   400
      Left            =   1920
      TabIndex        =   4
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
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
      Caption         =   "Hide"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton jcbutton1 
      Height          =   470
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   550
      _ExtentX        =   979
      _ExtentY        =   820
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
      Caption         =   ""
      Picture         =   "Form16.frx":113B0
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton jcbutton3 
      Height          =   400
      Left            =   2880
      TabIndex        =   6
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
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
      Caption         =   "exit"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   3915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3915
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
    Me.Show
End Sub

Private Sub Form_Load()
    refreshData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim msg As Long
    Dim sFilter As String
    
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN
            Me.Show
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormUtama.Show
End Sub

Private Sub jcbutton1_Click()
'    FormUtama.Show
    Unload Me
End Sub

Private Sub jcbutton2_Click()
    AddTray Me
End Sub

Private Sub jcbutton3_Click()
    tanya = MsgBox("Yakin Ingin Keluar ? ", vbYesNo + vbInformation, "Keluar dari Program")
    If tanya = vbYes Then
        On Error Resume Next
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = "Pesan Keluar [" & outBox.RecordCount & "]"
    Label2.Caption = "Pesan Terkirim [" & sentBox.RecordCount & "]"
    inBox.Filter = "Processed ='" & False & "'"
    Label3.Caption = "Pesan Masuk [" & inBox.RecordCount & "]"
End Sub

Private Sub Timer2_Timer()

Dim nmKKI, Krit, d As String
Dim i, j, k As Integer
Dim tSMS As String
Dim hpF As String
With xinBox
.Requery
    .Filter = "TextDecoded like 'jdw %' AND processed=false"
    If Not .EOF Then
        .MoveFirst
        'Timer2.Enabled = False
    '    For i = 1 To .RecordCount
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
            Text1.text = Text1.text & vbCrLf & Now & "- " & !Sendernumber & " cek jadwal pada halaqah " & nmKKI & ".."
            !Processed = True
            .MoveNext
    '    Next i
       ' Timer2.Enabled = True
    End If
    
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
            Text1.text = Text1.text & vbCrLf & Now & "- " & !Sendernumber & " cek data Murobbi pada halaqah " & nmKKI & ".."
            
            !Processed = True
            .MoveNext
        Next i
        'Timer2.Enabled = True
    End If
    
    
    .Filter = "TextDecoded like 'halkec %' AND processed=false"
    
    If Not .EOF Then
        .MoveFirst
        'Timer2.Enabled = False
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
                If Len(tSMS) <= 158 Then
                    kirimSMS tSMS, !Sendernumber
                Else
                    smsPart tSMS, !Sendernumber
                End If
            End If
            Text1.text = Text1.text & vbCrLf & Now & "- " & !Sendernumber & " cek halaqah pada wil " & Krit & ".."
            
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
            xrskki.Filter = "Kecamatan like '" & Krit & "' AND Tipe='Kampus'"
            
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
            Text1.text = Text1.text & vbCrLf & Now & "- " & !Sendernumber & " cek halaqah pada wil " & Krit & ".."
            
            !Processed = True
            .MoveNext
        Next i
    End If
    'Timer2.Enabled = True
    
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
            Text1.text = Text1.text & vbCrLf & Now & "- " & !Sendernumber & " sms ke ikh hal " & Krit & ".."
            
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
    
    cekupMur ("gtHr")
    
    cekupMur ("gtwkt")
    
    cekupMur ("gtT4")
    
.Requery
End With


End Sub
