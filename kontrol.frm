VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form menkontrol 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Список Шиндлера"
   ClientHeight    =   12645
   ClientLeft      =   3540
   ClientTop       =   1350
   ClientWidth     =   14400
   Icon            =   "kontrol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12645
   ScaleWidth      =   14400
   Begin VB.CommandButton cmdinfo 
      Caption         =   "Информация"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   11880
      Width           =   2415
   End
   Begin VB.CommandButton cmdlock 
      Caption         =   "Заблокировать"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   11160
      Width           =   2415
   End
   Begin VB.ComboBox ypk_vibor 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   10560
      Width           =   2415
   End
   Begin MSComctlLib.ListView listres 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   17383
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "№"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ФИО"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Г.Р."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ВК"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Куда"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Кто"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "#"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "B"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Label txtdatapr 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   9600
      TabIndex        =   11
      Top             =   12000
      Width           =   3495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Дата привоза"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9600
      TabIndex        =   10
      Top             =   11400
      Width           =   3495
   End
   Begin VB.Label txtvk 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   400
      Left            =   3360
      TabIndex        =   9
      Top             =   12000
      Width           =   3500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Прозрачно
      Caption         =   "ВК"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3360
      TabIndex        =   8
      Top             =   11400
      Width           =   3500
   End
   Begin VB.Label txtdatar 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   9600
      TabIndex        =   7
      Top             =   10680
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Дата рождения"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9600
      TabIndex        =   6
      Top             =   10080
      Width           =   3495
   End
   Begin VB.Label txtname 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3360
      TabIndex        =   5
      Top             =   10680
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Ф.И.О"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3360
      TabIndex        =   4
      Top             =   10080
      Width           =   3500
   End
   Begin VB.Menu to_excel 
      Caption         =   "В Excel"
   End
   Begin VB.Menu km_add 
      Caption         =   "Доб. призывника"
   End
   Begin VB.Menu mnuserv 
      Caption         =   "Сервис"
      Begin VB.Menu mm_edit 
         Caption         =   "Редактировать"
      End
      Begin VB.Menu line01 
         Caption         =   "-"
      End
      Begin VB.Menu disp_all 
         Caption         =   "Показать всех"
      End
      Begin VB.Menu disp_lock 
         Caption         =   "Показать заблокированных"
      End
      Begin VB.Menu show_ost 
         Caption         =   "Показать оставшихся"
      End
      Begin VB.Menu line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Удалить"
      End
   End
End
Attribute VB_Name = "menkontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim type_lock As Boolean
Private Sub CommAND1_Click()
gen_list (2)
End Sub
Public Sub gen_list(type_gen As Long)
On Error Resume Next
Dim x As Long
Dim Y As Long
Dim fiol As String


Dim gg As String
Dim z As Long
listres.ListItems.Clear
If type_gen = 0 Then Call mysql.query("SELECT id,fio,gr,vk,kyda,kto,lock_pr FROM control_" & nowBase & " ORDER BY vk,fio ASC")
If type_gen = 1 Then Call mysql.query("SELECT id,fio,gr,vk,kyda,kto,lock_pr FROM control_" & nowBase & " WHERE lock_pr='1' ORDER BY vk,fio ASC")
If type_gen = 2 Then Call mysql.query("SELECT id,fio,gr,vk,kyda,kto,lock_pr FROM control_" & nowBase & " WHERE lock_pr='0' ORDER BY vk,fio ASC")
Y = st
ReDim id_pr(st)
ReDim fio(st)
ReDim gr(st)
ReDim vk(st)
ReDim kyda(st)
ReDim kto(st)
ReDim lockpr(st)
    For x = 1 To st
        id_pr(x) = DAT(1, x)
        fio(x) = DAT(2, x)
        gr(x) = DAT(3, x)
        vk(x) = DAT(4, x)
        kyda(x) = DAT(5, x)
        kto(x) = DAT(6, x)
        lockpr(x) = DAT(7, x)
    Next x

For x = 1 To Y

        fiol = Left$(fio(x), 4)
        Set LF = listres.ListItems.Add()
        LF.SubItems(1) = id_pr(x)
        LF.SubItems(2) = fio(x)
        LF.SubItems(3) = gr(x)
        LF.SubItems(4) = vk(x)
        LF.SubItems(5) = kyda(x)
        LF.SubItems(6) = kto(x)
        mysql.query ("SELECT idprnik FROM prnik_" & nowBase & " WHERE fam like '" & fiol & "%' AND txtvk like '" & vk(x) & "%'")
            If st > 0 Then
                If lockpr(x) = 1 Then
                LF.ForeColor = &HFF&
                            For nC = 1 To 6
                                LF.ListSubItems.Item(nC).ForeColor = &HFF&
                            Next nC
                End If
                LF.SubItems(7) = st
            End If
            
            If lockpr(x) = 1 Then LF.SubItems(8) = "Да"
Next x
End Sub

Private Sub cmdinfo_Click()
On Error Resume Next
    expupk = ypk_vibOR.Text
    frmInfoPr.Show vbModal, Me
End Sub

Private Sub cmdlock_Click()
If type_lock = True Then
Call mysql.query("SELECT kyda,kto FROM control_" & nowBase & " WHERE id='" & id_mk & "'")
Dim kyda As String
Dim kto As String
kyda = DAT(1, 1)
kto = DAT(2, 1)

Call mysql.query("SELECT `lock` FROM prnik_" & nowBase & " WHERE `idprnik`='" & ypk_vibOR.Text & "'")
If DAT(1, 1) = "0" Then
Call mysql.query("UPDATE prnik_" & nowBase & " SET `lprim` = 'Список Шиндлера в " & kyda & ". Заблокировано " & kto & "' WHERE `idprnik` = " & ypk_vibOR.Text)
Call mysql.query("UPDATE prnik_" & nowBase & " SET `lock` = '3' WHERE `idprnik` = " & ypk_vibOR.Text)
End If
Call mysql.query("UPDATE `control_" & nowBase & "` set `lock_pr`='1', `lock_id`='" & ypk_vibOR.Text & "' WHERE `id`='" & id_mk & "'")

MsgBox "Итак мы заблокировали его!", vbExclamation, "Блокировка"
menkontrol.gen_list (2)
Else
Call mysql.query("SELECT `lock` FROM `prnik_" & nowBase & "` WHERE `idprnik`='" & ypk_vibOR.Text & "'")
If DAT(1, 1) = 3 Then
Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `lprim` = '' WHERE `idprnik` = " & ypk_vibOR.Text)
Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `lock` = '0' WHERE `idprnik` = " & ypk_vibOR.Text)
End If
Call mysql.query("UPDATE `control_" & nowBase & "` set `lock_pr`='0', `lock_id`='0' WHERE `id`='" & id_mk & "'")

MsgBox "Итак мы разблокировали его!", vbExclamation, "Блокировка"

menkontrol.gen_list (2)
End If

End Sub

Private Sub disp_all_Click()
gen_list (0)
End Sub

Private Sub disp_lock_Click()
gen_list (1)
End Sub
Private Sub listres_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
        If Button = 2 Then Call PopupMenu(mnuserv)
End Sub
Private Sub listres_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If KeyCode = 45 Then mk_add.Show vbModal, Me
  If KeyCode = 27 Then Unload Me
    End Sub
Private Sub FORm_Load()
Call Clear_Place
gen_list (2)
End Sub

Private Sub km_add_Click()
mk_add.Show vbModal, Me
End Sub
Private Sub Clear_Place()
ypk_vibOR.Clear
txtname = vbNullString
txtvk = vbNullString
txtdatar = vbNullString
txtdatapr = vbNullString
End Sub
Private Sub listres_Click()
On Error Resume Next
id_mk = listres.SelectedItem.SubItems(1)
Call Clear_Place
If Len(listres.SelectedItem.SubItems(7)) > 0 Then
    If Len(listres.SelectedItem.SubItems(8)) = 0 Then
               type_lock = True
          
          
            Dim fio As String
            Dim vk As String
            If ypk_vibOR.Enabled = True Then
            mysql.query ("SELECT fio,vk FROM control_" & nowBase & " WHERE id='" & id_mk & "'")
            fio = DAT(1, 1)
            fio = Left$(fio, 4)
            vk = DAT(2, 1)
            mysql.query ("SELECT idprnik FROM prnik_" & nowBase & " WHERE fam like '" & fio & "%' AND txtvk like '" & vk & "%'")
                For x = 1 To st
                    ypk_vibOR.AddItem (DAT(1, x))
                Next x
                    ypk_vibOR.Enabled = True
                End If
                ypk_vibOR.ListIndex = 0
                ypk_vibOR.Enabled = True
                cmdlock.Caption = "Заблокировать"

              
          
          
          Call mysql.query("SELECT fam,name,otch, txtvk,datar, dataosp FROM prnik_" & nowBase & " WHERE idprnik = '" & ypk_vibOR.Text & "'")
            txtname.Caption = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1)
            txtvk.Caption = DAT(4, 1)
            txtdatar.Caption = CnvDataSqLToWin(DAT(5, 1))
            txtdatapr.Caption = CnvDataSqLToWin(DAT(6, 1))
    Else
        type_lock = False
        mysql.query ("SELECT lock_id FROM control_" & nowBase & " WHERE id='" & id_mk & "'")
        ypk_vibOR.AddItem (DAT(1, 1))
        ypk_vibOR.ListIndex = 0
        ypk_vibOR.Enabled = False
        mysql.query ("SELECT fam,name,otch, txtvk,datar, dataosp FROM prnik_" & nowBase & " WHERE idprnik = '" & DAT(1, 1) & "'")
        If st > "0" Then
        txtname.Caption = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1)
        txtvk.Caption = DAT(4, 1)
        txtdatar.Caption = CnvDataSqLToWin(DAT(5, 1))
        txtdatapr.Caption = CnvDataSqLToWin(DAT(6, 1))
        cmdlock.Caption = "Разблокировать"
        End If
        
    End If
Else

    ypk_vibOR.Enabled = True
    cmdlock.Caption = "Заблокировать"
End If

'самозаполненение

'ddd



End Sub

Private Sub mk_edit_pr_Click()

End Sub

Private Sub mk_razblokirovatj_Click()

End Sub

Private Sub mm_edit_Click()
On Error Resume Next
id_mk = listres.SelectedItem.SubItems(1)
frmEditMK.Show vbModal, Me
End Sub
Private Sub listres_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    
    listres.Sorted = True
    
    If listres.SortKey = ColumnHeader.Index - 1 Then
        If listres.SortOrder = lvwDescending Then listres.SortOrder = lvwAscending Else listres.SortOrder = lvwDescending
    Else
        listres.SortOrder = lvwAscending
        listres.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub mnudel_Click()
id_mk = listres.SelectedItem.SubItems(1)
If MsgBox("Вы уверены, что хотите снять с контроля этого призывника?", vbOKCancel, "Удаление призывника " & listres.SelectedItem.SubItems(2)) = vbOK Then
mysql.query ("DELETE FROM control_" & nowBase & " WHERE id='" & id_mk & "'")
gg = MsgBox("Успешно удалили призывника", vbOKOnly, "Удаление")
End If


End Sub

Private Sub show_ost_Click()
gen_list (2)
End Sub

Private Sub to_excel_Click()
Call Cnv.ResoultSearch(listres, True, "ListCommAND", Caption)
End Sub

Private Sub ypk_vibOR_Click()
Dim fio As String
            Call mysql.query("SELECT fam,name,otch, txtvk,datar, dataosp FROM prnik_" & nowBase & " WHERE idprnik = '" & ypk_vibOR.Text & "'")
            txtname.Caption = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1)
            txtvk.Caption = DAT(4, 1)
            txtdatar.Caption = CnvDataSqLToWin(DAT(5, 1))
            txtdatapr.Caption = CnvDataSqLToWin(DAT(6, 1))
End Sub
