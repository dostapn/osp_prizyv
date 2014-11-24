VERSION 5.00
Begin VB.Form frmUpk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Добавить призывника в команду"
   ClientHeight    =   4275
   ClientLeft      =   5580
   ClientTop       =   930
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6840
   Visible         =   0   'False
   Begin VB.CheckBox Checkvod 
      Caption         =   "Водитель"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   330
      Left            =   5595
      TabIndex        =   8
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   150
      TabIndex        =   13
      Top             =   3600
      Width           =   6465
   End
   Begin VB.TextBox txtDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1575
      TabIndex        =   6
      Top             =   2340
      Width           =   2730
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   150
      TabIndex        =   10
      Top             =   1080
      Width           =   6465
   End
   Begin VB.PictureBox picAddPrnik 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   6810
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   6840
      Begin VB.Label lblVod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Если призывник является водителем, не забудте поставить галочку."
         Height          =   195
         Left            =   525
         TabIndex        =   12
         Top             =   555
         Width           =   5355
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Введите номер учетно-послужной карточки"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   300
         Width           =   3930
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6150
         Picture         =   "frmUpk.frx":08CA
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.TextBox txtVus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1575
      TabIndex        =   3
      Text            =   "Лин."
      Top             =   1965
      Width           =   2730
   End
   Begin VB.TextBox txtUpk 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1575
      TabIndex        =   1
      Top             =   1590
      Width           =   2730
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   4470
      TabIndex        =   7
      Top             =   3840
      Width           =   1005
   End
   Begin VB.ComboBox lstVus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3120
      Width           =   4950
   End
   Begin VB.Label dir_count 
      Caption         =   "Label6"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label ostatok_vrp 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5520
      TabIndex        =   18
      Top             =   3120
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Остаток"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Специальности"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label FIO 
      Caption         =   "Фамилия Имя Отчетво призывника"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Директива:"
      Height          =   195
      Left            =   525
      TabIndex        =   5
      Top             =   2325
      Width           =   900
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   300
      Left            =   1485
      Top             =   2295
      Width           =   2865
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   300
      Left            =   1485
      Top             =   1920
      Width           =   2865
   End
   Begin VB.Shape ShapeUpk 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   300
      Left            =   1485
      Top             =   1545
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&ВУС:"
      Height          =   195
      Left            =   525
      TabIndex        =   2
      Top             =   1965
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&УПК:"
      Height          =   210
      Left            =   525
      TabIndex        =   0
      Top             =   1590
      Width           =   375
   End
End
Attribute VB_Name = "frmUpk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Checkvod_click()

If Checkvod.VAlue = 1 Then txtVus.Text = "Вод. " Else txtVus.Text = "Лин."

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub CommAND1_Click()
On Error Resume Next
If Len(Trim(txtUpk)) = 0 Then Exit Sub
Dim lprim As String
Dim make_vus As String
On Error Resume Next

    INP_UPK = VAl(Trim(txtUpk))
    'txtDir = Trim$(txtDir)
       Dim c As Long
       Dim tCol As Long
       Dim L_VOD As Long
       Dim lngLock As Long

        L_VOD = Checkvod
   

           For c = 1 To frmCommand.listprnik.ListItems.Count
                If VAl(frmCommand.listprnik.ListItems(c).Text) = INP_UPK Then
                        tCol = frmCommand.listprnik.ListItems(c).ForeColor
                        frmCommand.listprnik.ListItems(c).EnsureVisible
                        frmCommand.listprnik.ListItems(c).Selected = True
                        frmCommand.listprnik.ListItems(c).Selected = False
                        Call SetTransparent(frmUpk.hWnd, 100)
                    MsgBox "Призывник уже в этой команде!", vbExclamation, strMAIN_TITLE
                        Call SetTransparent(frmUpk.hWnd, 255)
                        frmCommand.listprnik.ListItems(c).Selected = True
                        txtUpk.SetFocus
                        txtUpk.SelStart = 0: txtUpk.SelLength = Len(txtUpk.Text)
                    Exit Sub
                End If
            Next c
       
    

    Call mysql.query("SELECT `idprnik`, `lock` FROM prnik_" & nowBase & " WHERE idprnik = '" & INP_UPK & "'")
    If VAl(st) = 0 Then MsgBox "Такого призывника нет!", vbExclamation, txtUpk: txtUpk.SetFocus: txtUpk.SelStart = 0: txtUpk.SelLength = Len(txtUpk): Exit Sub
    
    If CLng(DAT(2, 1)) = 2 Then MsgBox "Простите, призывник находится в убывшей команде. Трогать его нельзя!", vbExclamation, strMAIN_TITLE: txtUpk.SetFocus: txtUpk.SelStart = 0: txtUpk.SelLength = Len(txtUpk.Text): Exit Sub


ret:
    Call mysql.query("SELECT prnik_" & nowBase & ".idprnik, naryad_" & nowBase & ".oblkom FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".idprnik = " & INP_UPK & " AND prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid;")
    

    
    If VAl(st) = 0 Then ' если призывник не в команде
 
            
            ' Получаем ФИО и Военкомат призывника по номеру УПК
            
                    Call mysql.query("SELECT `fam`, `name`, `otch` , `txtvk`, `lock`, `lprim` FROM prnik_" & nowBase & " WHERE `idprnik` = '" & INP_UPK & "'")
                    Dim lfio As String
                    Dim lVk As String
                    lfio = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1)
                    lVk = DAT(4, 1)
                    lprim = DAT(6, 1)
                    lngLock = DAT(5, 1)
                                        
                    'Обычная блокировка
                    If lngLock = "1" Then
                        If MsgBox("Этот призывник заблокирован.Блокировка:" & lprim & ". Вы уверены что этот призывник должен попасть в эту команду?", vbExclamation + vbOKCancel, "Система защиты") = vbOK Then
                            Resume
                        Else
                            Exit Sub
                        End If
                     End If
                    'конец обычной блокировки
                    If lngLock = "2" Then MsgBox "Этот призывник находится в отправленной команде.Вы должны разблокировать эту команду прежде чем удалять призывников из нее", vbExclamation + vbOKOnly, "Система защиты": Exit Sub
                    If lngLock = "3" Then
                       If Not acl = "G" Then Exit Sub
                        If Not MsgBox("Этот призывник заблокирован Администратором!!!Причина:" & lprim & ". Вы уверены что стоит его добавлять в эту команду?", vbCritical + vbOKCancel, "Система защиты") = vbOK Then Exit Sub
                    End If
                                              
                    
                    
                                        
                    Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `otprvid` = '" & p_com_id & "' WHERE `idprnik` = '" & INP_UPK & "'")
                    Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kolvo` = (SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid = " & p_com_id & ") WHERE otpravkaid = '" & p_com_id & "'")
                    Call mysql.query("SELECT naryad_" & nowBase & ".oblkom FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE prnik_" & nowBase & ".idprnik = " & INP_UPK & " AND prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid;")
                    Dim oblkom As String
                    oblkom = DAT(1, 1)
                    Call log_sql("0", "0", INP_UPK, oblkom)
                    Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vod` = '" & L_VOD & "' WHERE idprnik = '" & INP_UPK & "'")
                    
'make_vus = txtVus
'If txtVus = "Вод. С" Or txtVus = "Вод. C" Then make_vus = "837"
'If txtVus = "Вод. Д" Then make_vus = "845"
'If txtVus = "Вод. E" Or txtVus = "Вод. Е" Then make_vus = "846"
'If txtVus = "Вод. АПА" Then make_vus = "283"
'If txtVus = "Вод. ЭМ" Then make_vus = "837 ЭМ"

                    Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vus` = '" & txtVus & "' WHERE idprnik = '" & INP_UPK & "'")
                    Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `dir` = '" & txtDir & "' WHERE idprnik = '" & INP_UPK & "'")
                    
    Else   '  призывник уже в команде
            Dim tmpCom As String
            Dim OTPRv_ID As Long
            Dim tColVo As Long
            tmpCom = DAT(2, 1)

                     Call mysql.query("SELECT `fam`, `name`, `otch` , `txtvk`, `lock`, `lprim` FROM prnik_" & nowBase & " WHERE `idprnik` = '" & INP_UPK & "'")
                     lprim = DAT(6, 1)
                     lngLock = DAT(5, 1)
                     If lngLock = "1" Then
                        If MsgBox("Этот призывник заблокирован.Блокировка:" & lprim & "Вы уверены что этот призывник должен попасть в эту команду?", vbExclamation + vbOKCancel, "Система защиты") = vbYes Then
                            Resume
                        Else
                            Exit Sub
                        End If
                     End If
                    'конец обычной блокировки
                    If lngLock = "2" Then MsgBox "Этот призывник находится в отправленной команде.Вы должны разблокировать эту команду прежде чем удалять призывников из нее", vbExclamation + vbOKOnly, "Система защиты": Exit Sub
                    If lngLock = "3" Then
                       If Not acl = "G" Then Exit Sub
                        If Not MsgBox("Этот призывник заблокирован Администратором!!!Причина:" & lprim & ". Вы уверены что стоит его добавлять в эту команду?", vbCritical + vbOKCancel, "Система защиты") = vbOK Then Exit Sub
                    End If

                                                                

            If MsgBox("Призывник " & NL2 & DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1) & NL2 & "в команде " & tmpCom & ". Перезаписываем?", vbQuestion + vbOKCancel, "Подтверждение") = vbOK Then
                
                Dim old_id_c As String
                Call mysql.query("SELECT `otprvid` FROM prnik_" & nowBase & " WHERE idprnik='" & INP_UPK & "'")
                old_id_c = DAT(1, 1)
                
                Call frmCommand.del_sql(INP_UPK)
                Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `otprvid` = '" & p_com_id & "' WHERE `idprnik` = '" & INP_UPK & "'")
                
                Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vod` = '" & L_VOD & "' WHERE idprnik = " & INP_UPK)
                Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vus` = '" & txtVus & "' WHERE idprnik = " & INP_UPK)
                Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `dir` = '" & txtDir & "' WHERE idprnik = " & INP_UPK)
                
                Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kolvo` = (SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid='" & old_id_c & "') WHERE `otpravkaid` = '" & old_id_c & "'")
                Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kolvo` = (SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid='" & p_com_id & "') WHERE `otpravkaid` = '" & p_com_id & "'")
                
                Call mysql.query("SELECT oblkom FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE  naryad_" & nowBase & ".narid= otpravka_" & nowBase & ".narid AND otpravkaid='" & p_com_id & "'")
                Call log_sql("0", "0", INP_UPK, DAT(1, 1))
            End If

            
    End If
    
                txtUpk.SetFocus: txtUpk.SelStart = 0: txtUpk.SelLength = Len(txtUpk.Text)
                
                c = VAl(txtUpk)
                
                frmCommand.prnik_load (0)
                frmCommand.listprnik.Refresh
                frmCommand.commANDs_load
                frmCommand.listcommands.Refresh
                If Not INP_UPK = 0 Then frmCommand.listprnik.FindItem(c).Selected = True: frmCommand.listprnik.FindItem(c).EnsureVisible

frmCommand.listprnik.FindItem(c).EnsureVisible
frmCommand.listprnik.FindItem(c).ForeColor = &HFF&
frmCommand.listprnik.FindItem(c).Selected = True


End Sub
Function LockPrnik(lngUpk As Long) As Boolean
        LockPrnik = mysql.query("UPDATE `prnik_" & nowBase & "` SET `Lock` = 1 WHERE idprnik = " & lngUpk)
        Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `lprim` = '" & "AutoLock" & "' WHERE idprnik = " & lngUpk)
End Function

Private Sub FORm_ActiVAte()
On Error Resume Next
Dim x As Long
Dim defInd As Long
Dim okr As String
Dim vusc As String
Dim ost As Long
okr = frmCommand.listcommands.ListItems(frmCommand.listcommands.SelectedItem.Index).SubItems(8)
vusc = frmCommand.listcommands.ListItems(frmCommand.listcommands.SelectedItem.Index).SubItems(1)
Call mysql.query("select vrp, oblkom, major,minor from naryad_" & nowBase & " where okrkom = '" & okr & "' group by vrp")
For x = 1 To st
         lstVus.AddItem DAT(1, x)
Next x


dir_count.Caption = (Int(DAT(3, 1)) - Int(DAT(4, 1)))



lstVus.Text = vusc
Me.Visible = True
End Sub

Private Sub FORm_Initialize()
Me.Visible = False
End Sub


Private Sub lstVus_click()
Dim okr As String
Dim srez As String
okr = frmCommand.listcommands.ListItems(frmCommand.listcommands.SelectedItem.Index).SubItems(8)
'Call mysql.query("select vrp, oblkom, sum(naryad_" & nowBase & ".kolvo), sum(dolg), sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_" & nowBase & ", naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = naryad_" & nowBase & ".narid and okrkom = '" & okr & "' and vrp = '" & lstVus.Text & "'")
'If st = 0 Then
'Call mysql.query("SELECT vrp, oblkom, sum(kolvo) FROM naryad_" & nowBase & " WHERE okrkom = '" & okr & "' and vrp = '" & lstVus.Text & "'")
'ostatok_vrp.Caption = Int(DAT(3, 1))
'Else
Call mysql.query("select sum(naryad_" & nowBase & ".kolvo), sum(dolg), (select sum(naryad_srezki_" & nowBase & ".kolvo) from naryad_srezki_" & nowBase & " where naryad_srezki_" & nowBase & ".narid = naryad_" & nowBase & ".narid) from naryad_" & nowBase & " where naryad_" & nowBase & ".okrkom = '" & okr & "' and naryad_" & nowBase & ".vrp = '" & lstVus.Text & "'")
If DAT(3, 1) = "" Then srez = 0 Else srez = DAT(3, 1)
ostatok_vrp.Caption = Int(Int(DAT(1, 1)) + Int(DAT(2, 1)) - srez)
'End If

End Sub

Private Sub txtUpk_Change()
On Error Resume Next
        
        Call mysql.query("SELECT concat(fam,' ',name,' ',otch) FROM prnik_" & nowBase & " WHERE idprnik = '" & txtUpk & "'")
        If st = 0 Then
        FIO.Caption = "Не существует"
        Else
    FIO.Caption = DAT(1, 1)
        End If
End Sub

