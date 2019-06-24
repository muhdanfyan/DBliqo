VERSION 5.00
Begin VB.Form FSim 
   Caption         =   "Form16"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "FSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, b As String
rsKKI.MoveFirst
    For i = 1 To rsKKI.RecordCount
        b = rsKKI!murobbi
        
        rsKKI!murobbi = Right(b, Len(b) - 5)
        rsKKI.MoveNext
    Next i
End Sub

Private Sub Form_Load()
    masukDB
    rsIkh.Find "NamaL='RAHMAt'"
    
    If Not rsIkh.EOF Then
        Me.Caption = rsIkh!Hp
    End If
End Sub
