Attribute VB_Name = "Module1"
Public Konek As New ADODB.Connection
Public Kon As New ADODB.Connection
Public rsKKI As New ADODB.Recordset
Public rsIkh As New ADODB.Recordset
Public rsUst As New ADODB.Recordset
Public rsupIk As New ADODB.Recordset
Public rsupKKI As New ADODB.Recordset
Public aktif As New ADODB.Recordset
Public Taktif As New ADODB.Recordset
Public rsAkt As New ADODB.Recordset
Public outBox As New ADODB.Recordset
Public inBox As New ADODB.Recordset
Public xinBox As New ADODB.Recordset
Public sentBox As New ADODB.Recordset
Public user As New ADODB.Recordset
Public iB As New ADODB.Recordset
Public Baru As Boolean
Public ddn As Object
Public xrskki As New ADODB.Recordset
Public xrsust As New ADODB.Recordset
Public xrsIkh As New ADODB.Recordset
Public rsKonf As New ADODB.Recordset
Public req As Boolean

Public Sub masukDB()
    Konek.CursorLocation = adUseClient
    Kon.CursorLocation = adUseClient
    Kon.Open "sms"
    Konek.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Db.mdb;Persist Security Info=False"
    rsKKI.Open "Select * from dbKKI ORDER BY NamaKKI ASC", Konek, adOpenDynamic, adLockOptimistic
    rsIkh.Open "Select * from dbIkhwah ORDER BY NamaL ASC", Konek, adOpenDynamic, adLockOptimistic
    rsUst.Open "Select * from dbUst ORDER BY Nama ASC", Konek, adOpenDynamic, adLockOptimistic
    rsupIk.Open "Select * From UpIkh ORDER BY wkt DESC, KKI ASC", Konek, adOpenDynamic, adLockOptimistic
    aktif.Open "Select * From UpIkh ORDER BY wkt DESC, KKI ASC", Konek, adOpenDynamic, adLockOptimistic
    Taktif.Open "Select * From UpIkh ORDER BY wkt DESC, KKI ASC", Konek, adOpenDynamic, adLockOptimistic
    rsupKKI.Open "Select * from UpKKI ORDER BY wkt DESC, NamaKKI ASC", Konek, adOpenDynamic, adLockOptimistic
    rsAkt.Open "SELECT DbKKI.NamaKKI, DbKKI.Murobbi, DbKKI.Naqib,UpKKI.aktif, UpKKI.Taktif, DbKKI.Tempat, DbKKI.Hari, DbKKI.Waktu, DbKKI.Jenis," & _
    "DbKKI.Tahun, DbKKI.Tipe,  UpKKI.wkt, UpKKI.Ket  From DbKKI, UpKKI WHERE DbKKI.NamaKKI=UpKKI.NamaKKI" _
    , Konek, adOpenDynamic, adLockOptimistic
    
    outBox.Open "Select UpdatedInDB,DestinationNumber,TextDecoded from outbox ORDER BY UpdatedInDB DESC", Kon, adOpenDynamic, adLockOptimistic
    inBox.Open "Select ReceivingDateTime,SenderNumber,TextDecoded, processed from inbox ORDER BY Processed ASC,ReceivingDateTime DESC", Kon, adOpenDynamic, adLockOptimistic
    xinBox.Open "Select ReceivingDateTime,SenderNumber,TextDecoded, processed from inbox ORDER BY Processed ASC,ReceivingDateTime DESC", Kon, adOpenDynamic, adLockOptimistic
    iB.Open "Select ReceivingDateTime,SenderNumber,TextDecoded, processed from inbox ORDER BY Processed ASC,ReceivingDateTime DESC", Kon, adOpenDynamic, adLockOptimistic
    sentBox.Open "Select UpdatedInDB,DestinationNumber,TextDecoded from sentitems ORDER BY UpdatedInDB DESC", Kon, adOpenDynamic, adLockOptimistic
    user.Open "Select * from tbuser ORDER BY user ASC", Konek, adOpenDynamic, adLockOptimistic
    rsKonf.Open "Select * from tbKeg ORDER BY tgl DESC", Konek, adOpenDynamic, adLockOptimistic
    
    xrskki.Open "Select * from dbKKI ORDER BY NamaKKI ASC", Konek, adOpenDynamic, adLockOptimistic
    xrsust.Open "Select * from dbUst ORDER BY Nama ASC", Konek, adOpenDynamic, adLockOptimistic
    xrsIkh.Open "Select * from dbIkhwah ORDER BY NamaL ASC", Konek, adOpenDynamic, adLockOptimistic
    
End Sub

Public Sub kirimSMS(text As String, noHp As String)
    With outBox
        .AddNew
        !Textdecoded = text
        !DestinationNumber = noHp
        .Update
    End With
End Sub

Public Sub smsPart(txSMS As String, nhp As String)
Dim i, j, k, l, n, m As Integer

    k = Int(Len(txSMS) / 152 + 1)
    n = 1
    For j = 1 To k
        If j = k Then
            l = k Mod 152
        Else
            l = 152
        End If
            kirimSMS "info" & j & "/" & k & ":" & Mid(txSMS, n, l), nhp
        n = n + 152
    Next j

End Sub


Public Sub sms(t As String, HpX As String)

If HpX <> "" Or HpX <> "-" Then
    If Len(t) <= 158 Then
        kirimSMS t, HpX
    Else
        smsPart t, HpX
    End If
End If
End Sub

Public Sub admin(a As Boolean)
    With MDI
        .Datkti.Visible = a
        .datUst.Visible = a
        .login.Visible = Not a
        .user.Visible = a
        .sms.Visible = a
    End With
End Sub

Public Sub refreshData()
    rsIkh.Requery
    rsUst.Requery
    rsupIk.Requery
    aktif.Requery
    Taktif.Requery
    rsupKKI.Requery
    rsAkt.Requery
    outBox.Requery
    inBox.Requery
    xinBox.Requery
    iB.Requery
    sentBox.Requery
    user.Requery
End Sub

Public Sub inputKonf(nmX As String, tglX As Date, HpX As String, kkiX As String, konX As String, drX As String)
    With rsKonf
        .Find "hp='" & HpX & "'"
        If .EOF Then .AddNew
        !nama = nmX
        !tgl = tglX
        !Hp = HpX
        !KKI = kkiX
        !konf = konX
        !dari = drX
        .Update
    End With
End Sub

Public Sub cekupMur(t As String)
Dim hpF As String

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
        
        MsgBox "[" & Krit & "]"
        
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
                Form16.Text1.text = Form16.Text1.text & vbCrLf & Now & "-" & !Sendernumber & " mengupdate data hal " & Krit
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


Public Sub ckNmr(nhp As String, hpb As String)
    
    If Mid(nhp, 1, 3) = "+62" Then
        hpb = "0" & Right(nhp, Len(nhp) - Len("+62"))
    Else
        hpb = nhp
    End If
    
End Sub

