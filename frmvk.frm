VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnaryad_info 
   Caption         =   "Информация о команде"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8235
   Icon            =   "frmvk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Долг"
      Height          =   1575
      Left            =   5160
      TabIndex        =   28
      Top             =   8160
      Width           =   3015
      Begin VB.CommandButton cmd_dolg 
         Caption         =   "Изменить"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtdolg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Text            =   "0"
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Срезки"
      Height          =   3015
      Left            =   5160
      TabIndex        =   25
      Top             =   5040
      Width           =   3015
      Begin VB.CommandButton cmdsrez 
         Caption         =   "Срезать"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtkolvosrez 
         Height          =   405
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker txtdatesrez 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16121857
         CurrentDate     =   39260
      End
      Begin VB.Label Label16 
         Caption         =   "Количество"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Дата"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Отправка"
      Height          =   4935
      Left            =   5160
      TabIndex        =   20
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.ListView listres 
         Height          =   4095
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7223
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "id"
            Object.Width           =   4
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Дата"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Кол-во"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label txtotp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Всего"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   4440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton opt3 
         Caption         =   "ЗАТО"
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   8400
         Width           =   1335
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Подшефная часть"
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   8400
         Width           =   1815
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Обычная"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   8400
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox txtokr 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   36
         Text            =   "txtokr"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox txtpodrod 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   37
         Text            =   "txtpodrod"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox txtdoroga 
         Height          =   315
         ItemData        =   "frmvk.frx":08CA
         Left            =   2520
         List            =   "frmvk.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   35
         Text            =   "txtdoroga"
         Top             =   3480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker txtdatanar 
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   5520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16121857
         CurrentDate     =   41233
      End
      Begin VB.TextBox txtprim 
         Height          =   2055
         Left            =   120
         TabIndex        =   24
         Top             =   6240
         Width           =   4815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Отмена"
         Height          =   495
         Left            =   2400
         TabIndex        =   19
         Top             =   9120
         Width           =   1695
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Сохранить"
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   9120
         Width           =   1455
      End
      Begin VB.TextBox txtkolvo 
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox txtpred 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtvch 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtpunkt 
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtoblkom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtokrkom 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Подрод войск"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Примечание:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Дата наряда"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   5565
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Количество"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Предназначение"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Воинская часть"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Дорога"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Пункт"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Военный округ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Обл. ком."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Окр. ком."
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Menu mnu_otpravka 
      Caption         =   "Отправки"
      Begin VB.Menu mnu_edit 
         Caption         =   "Редактировать"
      End
      Begin VB.Menu mnu_del_srez 
         Caption         =   "Удалить срезку"
      End
   End
End
Attribute VB_Name = "frmnaryad_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_dolg_Click()
On Error Resume Next
If Len(txtdolg.Text) Then
    If MsgBox("Вы уверены, что хотите изменить долг?", vbInformation + vbOKCancel) = vbOK Then
        Call mysql.query("UPDATE naryad_" & nowBase & " set dolg='" & txtdolg.Text & "' WHERE narid='" & txtid & "'")
        MsgBox "Вы удачно измнили долг", vbInformation
    End If
End If
End Sub

Private Sub cmd_save_Click()
On Error Resume Next
Dim in_com As String
Dim datt() As String
Dim okr_db As String
Dim okr() As String
Dim vch_type As Integer
okr = Split(txtokrkom, "/")
okr_e = okr(UBound(okr()))
If UBound(okr) - 1 > 0 Then
    okr_db = okr(0) & "/" & okr(1) & "/"
Else
    okr_db = okr(0) & "/"
End If

If opt1 = True Then vch_type = "0"
If opt2 = True Then vch_type = "1"
If opt3 = True Then vch_type = "2"
                
                
If cmd_save.Caption = "Добавить" Then
Call mysql.query("insert into naryad_" & nowBase & " set `okrkom` = '" & okr_db & "',`okrkom_e` = '" & okr_e & "',`oblkom` = '" & txtoblkom.Text & "',`rodv` = '" & txtpodrod.Text & "',`okr` = '" & txtokr.Text & "',`punkt` = '" & txtpunkt.Text & "',`doroga` = '" & txtdoroga.Text & "',`vch` = '" & txtvch.Text & "',`vrp` = '" & txtpred.Text & "',`kolvo` = '" & txtkolvo.Text & "',`datanar` = '" & CnvDataWinToSql(txtdatanar) & "',`dolg` = '" & txtdolg & "',`type` = '" & vch_type & "',`prim` = '" & txtprim & "'")
Else
Call mysql.query("update naryad_" & nowBase & " set `okrkom` = '" & okr_db & "',`okrkom_e` = '" & okr_e & "',`oblkom` = '" & txtoblkom.Text & "',`rodv` = '" & txtpodrod.Text & "',`okr` = '" & txtokr.Text & "',`punkt` = '" & txtpunkt.Text & "',`doroga` = '" & txtdoroga.Text & "',`vch` = '" & txtvch.Text & "',`vrp` = '" & txtpred.Text & "',`kolvo` = '" & txtkolvo.Text & "',`datanar` = '" & CnvDataWinToSql(txtdatanar) & "',`dolg` = '" & txtdolg & "',`type` = '" & vch_type & "',`prim` = '" & txtprim & "' WHERE narid='" & txtid & "'")
End If
'Call frmNaryad.rodvtr_Click
Unload Me
End Sub

Private Sub cmdsrez_Click()
On Error Resume Next
Dim newid As Long
Call mysql.query("SELECT `id` FROM naryad_srezki_" & nowBase & " where `narid`='" & txtid & "' and `data`='" & CnvDataWinToSql(txtdatesrez) & "'")
If VAl(st) = 0 Then
    Call mysql.query("SELECT max(id) FROM naryad_srezki_" & nowBase)
        If DAT(1, 1) = "" Then
            newid = 1
        Else
            newid = DAT(1, 1) + 1
        End If
    Call mysql.query("insert into naryad_srezki_" & nowBase & " VAlues ('" & newid & "','" & CnvDataWinToSql(txtdatesrez) & "','" & txtid & "','" & txtkolvosrez & "')")
    txtkolvosrez = ""
    Call info_otp
Else
    Call mysql.query("SELECT `kolvo` from naryad_srezki_" & nowBase & " where `narid`='" & txtid & "' and `data`='" & CnvDataWinToSql(txtdatesrez) & "'")
    Call mysql.query("UPDATE naryad_srezki_" & nowBase & " set kolvo='" & VAl(DAT(1, 1)) + VAl(txtkolvosrez) & "' where `narid`='" & txtid & "' and `data`='" & CnvDataWinToSql(txtdatesrez) & "'")
    Call info_otp
End If
End Sub
Private Sub CommAND2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim datt() As String
Dim x As Long

Call mysql.query("SELECT doroga FROM naryad_" & nowBase & " group by doroga")
For x = 1 To st
txtdoroga.AddItem (DAT(1, x))
Next x
Call mysql.query("SELECT `okr` FROM naryad_" & nowBase & " group by `okr`")
For x = 1 To st
     txtokr.AddItem (DAT(1, x))
Next x


If cmd_save.Caption = "Добавить" Then GoTo add
Call mysql.query("SELECT `narid`,`okrkom`,`okrkom_e`,`oblkom`,`rodv`,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`,`datanar`,`dolg`,`type`,`prim` FROM naryad_" & nowBase & " WHERE narid=" & p_com_id)
datt() = DAT()
txtid = datt(1, 1)
txtokrkom = datt(2, 1) & datt(3, 1)
txtoblkom = datt(4, 1)

Call mysql.query("SELECT `rodv` FROM naryad_" & nowBase & " WHERE `okr` = '" & datt(6, 1) & "' group by `rodv`")
For x = 1 To st
    txtpodrod.AddItem (DAT(1, x))
Next x



txtokr = datt(6, 1)
txtpunkt = datt(7, 1)
txtdoroga = datt(8, 1)
txtvch = datt(9, 1)
txtpred = datt(10, 1)
txtkolvo = datt(11, 1)
txtdatanar = CnvDataSqLToWin(datt(12, 1))
txtdolg = Int(datt(13, 1))
If datt(16, 1) = 0 Then opt1 = True
If datt(16, 1) = 1 Then opt2 = True
If datt(14, 1) = 2 Then opt3 = True
txtprim = datt(15, 1)
txtpodrod = datt(5, 1)

Call mysql.query("SELECT sum(kolvo) FROM naryad_srezki_" & nowBase & " WHERE `narid`='" & txtid & "'")
txtotp = DAT(1, 1)
Frame1.Caption = txtrodv & " " & txtpunkt
txtdatesrez = Date
Call info_otp
add:
End Sub
Private Sub info_otp()
listres.ListItems.Clear
Call mysql.query("SELECT `id`,`data`,`kolvo` FROM naryad_srezki_" & nowBase & " WHERE `narid` = '" & txtid & "'")
    For x = 1 To st
        Set LF = listres.ListItems.add()
        LF.SubItems(1) = DAT(1, x)
        LF.SubItems(2) = CnvDataSqLToWin(DAT(2, x))
        LF.SubItems(3) = DAT(3, x)
    Next x
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Call frmNaryad.rodvtr_Click
frmNaryad.bar_refresh
End Sub
Private Sub listres_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 2 Then Call PopupMenu(mnu_otpravka)
End Sub
Private Sub mnu_del_srez_Click()
If MsgBox("Вы уверены что хотите удалить данную срезку?", vbQuestion + vbYesNo, "Внимание") = vbYes Then
    Call mysql.query("DELETE FROM naryad_srezki_" & nowBase & " WHERE id='" & Int(listres.ListItems(listres.SelectedItem.Index).SubItems(1)) & "'")
    MsgBox "Вы удачно удалили срезку!", vbInformation
    Call info_otp
End If
End Sub
Private Sub mnu_edit_Click()
Dim srez As String
srez = InputBox("Изменение количество призывников отправленных " & listres.ListItems(listres.SelectedItem.Index).SubItems(2) & ":", "Изменение срезок.", Int(listres.ListItems(listres.SelectedItem.Index).SubItems(3)))
If Len(srez) > 0 Then Call mysql.query("UPDATE naryad_srezki_" & nowBase & " set kolvo='" & Int(srez) & "' WHERE id='" & Int(listres.ListItems(listres.SelectedItem.Index).SubItems(1)) & "'"): Call info_otp
End Sub


Private Sub txtkolvosrez_Change()
If Len(txtkolvosrez) > 0 Then
    cmdsrez.Enabled = True
    Else
    cmdsrez.Enabled = False
End If
End Sub

Private Sub txtokr_click()
On Error Resume Next
txtpodrod.Clear
Call mysql.query("SELECT rodv FROM naryad_" & nowBase & " WHERE okr = '" & txtokr.Text & "' group by rodv")
For x = 1 To st
    txtpodrod.AddItem (DAT(1, x))
Next x
End Sub
