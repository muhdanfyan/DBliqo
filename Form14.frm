VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form14 
   BorderStyle     =   0  'None
   Caption         =   "KIRIM SMS"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10890
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CKonf 
      Caption         =   "Konfirmasi"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   4545
   End
   Begin VB.Frame Frame2 
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   7200
      TabIndex        =   21
      Top             =   1320
      Width           =   5895
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   2640
         Width           =   2175
      End
      Begin VB.ListBox List1 
         Height          =   2985
         ItemData        =   "Form14.frx":10E2C
         Left            =   120
         List            =   "Form14.frx":10E2E
         TabIndex        =   22
         Top             =   120
         Width           =   2415
      End
      Begin DBLiqo.jcbutton btH 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "Hapus"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton btM 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   2640
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "Masukkan"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   6360
      TabIndex        =   17
      Top             =   600
      Width           =   5895
      Begin VB.CheckBox C3 
         Caption         =   "TAKWINIYAH"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox C2 
         Caption         =   "TA'ARIFIYAH"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox C1 
         Caption         =   "PRA TA'ARIFIYAH"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Value           =   1  'Checked
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   5895
      Begin VB.Frame FT 
         Caption         =   "HALAQOH"
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   3135
         Begin VB.CheckBox ckt3 
            Caption         =   "TAKWINIYAH"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox ckt2 
            Caption         =   "TA'ARIFIYAH"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   2655
         End
         Begin VB.CheckBox ckt1 
            Caption         =   "PRA TA'ARIFIYAH"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Murobbi"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Bendahara"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Naqib"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4545
      Begin VB.TextBox Text2 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Text            =   "WI-MKS"
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   3735
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SMS Massal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SMS Naqib dan Bendahara"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SMS Perorang"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin DBLiqo.jcbutton button1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "KIRIM"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton button2 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "BATAL"
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nhp As String

Private Sub btH_Click()
If List1.ListIndex >= 0 Then List1.RemoveItem List1.ListIndex
End Sub

Private Sub btM_Click()
    List1.AddItem Text3.text
    Text3.text = ""
End Sub

Sub smsJn(Jen As String, Tx As String)
    rsIkh.Filter = "Tingkat='" & Jen & "'"
        For i = 1 To rsIkh.RecordCount
            sms Tx, rsIkh!Hp
            If CKonf.Value = 1 Then inputKonf Tx, Now, rsIkh!Hp, rsIkh!NmHalaqa, "-", "dpc"
            rsIkh.MoveNext
        Next i
End Sub


Sub smsNq(Jen As String, Tx As String)
    rsKKI.Filter = "Jenis='" & Jen & "'"
        For i = 1 To rsKKI.RecordCount
            rsIkh.Filter = "NamaL='" & rsKKI!Naqib & "'"
            If Not rsIkh.EOF Then
                sms Tx, rsIkh!Hp
                If CKonf.Value = 1 Then inputKonf Tx, Now, rsIkh!Hp, rsKKI!namaKKI, "-", "dpc"
            End If
        rsKKI.MoveNext
        Next i
End Sub

Sub smsBd(Jen As String, Tx As String)
    rsKKI.Filter = "Jenis='" & Jen & "'"
        For i = 1 To rsKKI.RecordCount
            rsIkh.Filter = "NamaL='" & rsKKI!Bendahara & "'"
            If Not rsIkh.EOF Then
                sms Tx, rsIkh!Hp
                If CKonf.Value = 1 Then inputKonf Tx, Now, rsIkh!Hp, rsKKI!namaKKI, "-", "dpc"
            End If
        rsKKI.MoveNext
        Next i
End Sub

Sub smsMur(Jen As String, Tx As String)
    rsKKI.Filter = "Jenis='" & Jen & "'"
    For i = 1 To rsKKI.RecordCount
            rsUst.Filter = "Nama='" & rsKKI!murobbi & "'"
            If Not rsUst.EOF Then
                sms Tx, rsUst!Hp
                If CKonf.Value = 1 Then inputKonf Tx, Now, rsUst!Hp, rsKKI!namaKKI, "-", "dpc"
            End If
        rsKKI.MoveNext
    Next i
End Sub

Private Sub button1_Click()
Dim i As Integer, text As String
Dim tkt As String

If Text2.text <> "" Then
        text = Text1.text & Text2.text
    
 If CKonf.Value = 1 Then
        
    With rsKonf
    
    .Filter = "dari='dpc'"
    For i = 1 To rsKonf.RecordCount
        .Delete
        .Requery
    Next i
    
    End With
        
        text = text & " Konfirm[y/t]"
 End If

If Tab1.Tabs(1).Selected Then
    If C1.Value = 1 Then smsJn "PRA TAARIF", text
    If C2.Value = 1 Then smsJn "TAARIFIYAH", text
    If C3.Value = 1 Then smsJn "TAKWINIYAH", text
ElseIf Tab1.Tabs(2).Selected Then
    If ckt1.Value = 1 Then
        If Check1.Value = 1 Then smsNq "PRA TAARIF", text
        If Check2.Value = 1 Then smsBd "PRA TAARIF", text
        If Check3.Value = 1 Then smsMur "PRA TAARIF", text
    End If
    
    If ckt2.Value = 1 Then
        If Check1.Value = 1 Then smsNq "TAARIFIYAH", text
        If Check2.Value = 1 Then smsBd "TAARIFIYAH", text
        If Check3.Value = 1 Then smsMur "TAARIFIYAH", text
    End If
    
    If ckt3.Value = 1 Then
        If Check1.Value = 1 Then smsNq "TAKWINIYAH", text
        If Check2.Value = 1 Then smsBd "TAKWINIYAH", text
        If Check3.Value = 1 Then smsMur "TAKWINIYAH", text
    End If
    
    Dim em As String
    
ElseIf Tab1.Tabs(3).Selected Then
    
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        sms text, List1.text
        ckNmr List1.text, nhp
        rsIkh.Find "hp='" & nhp & "'"
        If rsIkh.EOF Then
            em = "-"
        Else
            em = rsIkh!NmHalaqa
        End If
        If CKonf.Value = 1 Then inputKonf text, Now, nhp, em, "-", "dpc"
        
    Next i
End If
Unload Me
End If
End Sub

Private Sub button2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    rsKKI.Filter = ""
    rsIkh.Filter = ""
End Sub

Private Sub Tab1_Click()
    If Tab1.Tabs(1).Selected Then
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = True
        Frame4.Top = Frame3.Top
        Frame4.Left = Frame3.Left
    ElseIf Tab1.Tabs(2).Selected Then
        Frame2.Visible = False
        Frame3.Visible = True
        Frame4.Visible = False
    ElseIf Tab1.Tabs(3).Selected Then
        Frame2.Visible = True
        Frame3.Visible = False
        Frame4.Visible = False
        Frame2.Top = Frame3.Top
        Frame2.Left = Frame3.Left
    End If
End Sub
