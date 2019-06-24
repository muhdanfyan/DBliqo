VERSION 5.00
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Aplikasi Database Halaqoh Tarbiyah"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "[&File]"
      Begin VB.Menu login 
         Caption         =   "&Login User"
      End
      Begin VB.Menu user 
         Caption         =   "&User"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Datkti 
      Caption         =   "[Data &KTI dan Data Kader]"
      Begin VB.Menu tampilKti 
         Caption         =   "&Tampilkan Data KTI"
      End
      Begin VB.Menu tampilIkh 
         Caption         =   "Tampilkan Data &Kader"
      End
      Begin VB.Menu AktIkh 
         Caption         =   "Tampilkan Ke&aktifan Kader"
      End
      Begin VB.Menu inputkti 
         Caption         =   "&Input Data KTI"
      End
      Begin VB.Menu aktifkti 
         Caption         =   "&Keaktifan KTI"
      End
   End
   Begin VB.Menu datUst 
      Caption         =   "[Data &Pembina/Ustadz Murobbi]"
      Begin VB.Menu tampilUst 
         Caption         =   "&Tampilkan Data Ustadz"
      End
      Begin VB.Menu IputUst 
         Caption         =   "&Input Data Ustadz"
      End
   End
   Begin VB.Menu sms 
      Caption         =   "[&SMS KADER]"
      Begin VB.Menu pesan 
         Caption         =   "&Lihat Pesan"
      End
      Begin VB.Menu kirim 
         Caption         =   "&Kirim SMS"
      End
      Begin VB.Menu konSMS 
         Caption         =   "Lihat K&onfirmasi SMS"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktifkti_Click()
    Form12.Show
End Sub

Private Sub AktIkh_Click()
    Form5.Show
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub inputkti_Click()
    Form2.Show
End Sub

Private Sub IputUst_Click()
    Form7.Show
End Sub

Private Sub kirim_Click()
   With Form14
    .Show
    .Tab1.Tabs(1).Selected = True
    End With
End Sub

Private Sub konSMS_Click()
Dim a1, a2, a3 As Integer

rsKonf.Filter = "dari='dpc' AND konf='y'"
a1 = rsKonf.RecordCount
rsKonf.Filter = "dari='dpc' AND konf='t'"
a2 = rsKonf.RecordCount
rsKonf.Filter = "dari='dpc' AND konf='-'"
a3 = rsKonf.RecordCount

    With FormUtama
        .Label4.Caption = a1 & " Ikhwah yang konfirmasi akan datang"
        .Label5.Caption = a2 & " Ikhwah yang konfirmasi akan tidak datang"
        .Label6.Caption = a3 & " Ikhwah yang belum konfirmasi"
        .Frame1.Visible = True
    End With
rsKonf.Filter = ""

End Sub

Private Sub login_Click()
    FLogin.Width = 4500
    FLogin.Height = 3000
    FLogin.Show
End Sub

Private Sub MDIForm_Load()
    masukDB
'    FormUtama.Show
End Sub

Private Sub tes_Click()
'    FormUtama.Show
End Sub

Private Sub pesan_Click()
    Form13.Show
End Sub

Private Sub tampilIkh_Click()
    Form3.Show
End Sub

Private Sub tampilKti_Click()
    Form1.Show
End Sub

Private Sub tampilUst_Click()
    Form8.Show
End Sub
