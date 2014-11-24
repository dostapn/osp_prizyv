VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmdir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Директивщики"
   ClientHeight    =   10215
   ClientLeft      =   1605
   ClientTop       =   2745
   ClientWidth     =   13980
   Icon            =   "frmdir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   13980
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Привязать вниз
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9960
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstprnik 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial CYR"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   4
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Фамилия"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Имя"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Отчество"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "ВК"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Г.Р."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "#"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Блокировка"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Примечание"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstdirs 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8916
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial CYR"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   4
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Директива"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Дата"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "В/Ч"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Кол-во"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Обл. ком."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Дата команды"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnu_directivi 
      Caption         =   "Директивы"
      Begin VB.Menu mnu_directivi_add 
         Caption         =   "Добавить"
      End
      Begin VB.Menu mnu_directivi_edit 
         Caption         =   "Редактировать"
      End
      Begin VB.Menu mnu_directivi_del 
         Caption         =   "Удалить"
      End
   End
   Begin VB.Menu mnu_directivwiki 
      Caption         =   "Директивщики"
      Begin VB.Menu mnu_dir_p_add 
         Caption         =   "Добавить"
      End
      Begin VB.Menu mnu_dir_p_edit 
         Caption         =   "Редактировать"
      End
      Begin VB.Menu mnu_dir_p_del 
         Caption         =   "Удалить"
      End
      Begin VB.Menu line01 
         Caption         =   "-"
      End
      Begin VB.Menu search 
         Caption         =   "Искать"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnu_print 
      Caption         =   "Печать"
      Begin VB.Menu mnu_print_vk 
         Caption         =   "Распределение по военкоматам"
      End
      Begin VB.Menu mnu_print_rodv 
         Caption         =   "Распределение по родам войск"
      End
   End
   Begin VB.Menu mnu_refresh 
      Caption         =   "Обновить"
   End
End
Attribute VB_Name = "frmdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Call sum_dir
Call load_dirs
Call view_nonblock
End Sub
Public Sub sum_dir()
Call mysql.query("select `id` from directivi_" & nowBase)
stt = st
datt() = DAT()
For x = 1 To stt
    Call mysql.query("select count(id) from directivi_p_" & nowBase & " where `did`='" & datt(1, x) & "'")
    Call mysql.query("update directivi_" & nowBase & " set `kolvo`='" & DAT(1, 1) & "' where `id`='" & datt(1, x) & "'")
Next x

End Sub
Public Sub load_dirs()
On Error Resume Next
Call mysql.query("select `id`, `dir`,`data`,`vch`,`kolvo`,`kom`,`kom_data` from directivi_" & nowBase)
lstdirs.ListItems.Clear
For x = 1 To st
    DAT(3, x) = CnvDataSqLToWin(DAT(3, x))
    DAT(7, x) = CnvDataSqLToWin(DAT(7, x))
    Set LF = lstdirs.ListItems.add(, , DAT(1, x))
    For c = 1 To 6
        LF.SubItems(c) = DAT(c + 1, x)
    Next c
Next x
lstdirs.Refresh
    Call ReSizeColumnHeaders(lstdirs)
    Call lstdirs_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call sum_dir
End Sub

Public Sub lstdirs_Click()
On Error Resume Next

Call mysql.query("select `id`,`fam`,`name`,`otch`,`vk`,`year_b`,`prim`,`pid` from `directivi_p_" & nowBase & "` where `did`='" & lstdirs.SelectedItem.Text & "'")
lstprnik.ListItems.Clear
For x = 1 To st
Set LF = lstprnik.ListItems.add(, , DAT(1, x))
    For c = 1 To 5
    LF.SubItems(c) = DAT(c + 1, x)
    Next c
LF.SubItems(8) = DAT(7, x)
If val(DAT(8, x)) > 0 Then LF.SubItems(7) = "ДА"
Next x


Call ReSizeColumnHeaders(lstprnik)
Call check_dirs_p
lstprnik.Refresh

End Sub

Private Sub lstdirs_DblClick()
frmdir_dirs.edit_pr = True: frmdir_dirs.Show vbModal, Me
End Sub

Private Sub lstprnik_DblClick()
frmdir_prnik.type_load = "1": frmdir_prnik.Show vbModal, Me
End Sub

Private Sub lstprnik_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then If acl = "G" Or acl = "O" Then frmdir_prnik.type_load = "0": frmdir_prnik.Show vbModal, Me
If KeyCode = 46 Then
    If acl = "G" Or acl = "O" Then
        If MsgBox("Вы уверены, что хотите удалить данного директивщика???", vbQuestion + vbYesNo, "Удаление") = vbYes Then
            Call mysql.query("DELETE FROM directivi_p_" & nowBase & " where `id`='" & lstprnik.SelectedItem.Text & "'")
        End If
    End If
End If
Call lstdirs_Click
End Sub

Private Sub lstprnik_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = "2" Then Call PopupMenu(mnu_directivwiki)
End Sub

Private Sub mnu_dir_p_add_Click()
Call lstprnik_KeyDown("45", "0")
End Sub

Private Sub mnu_dir_p_del_Click()
Call lstprnik_KeyDown("46", "0")
End Sub

Private Sub mnu_dir_p_edit_Click()
Call lstprnik_DblClick
End Sub

Private Sub mnu_directivi_add_Click()
frmdir_dirs.edit_pr = False: frmdir_dirs.Show vbModal, Me
End Sub

Private Sub mnu_directivi_edit_Click()
frmdir_dirs.edit_pr = True: frmdir_dirs.Show vbModal, Me
End Sub

Private Sub mnu_print_rodv_Click()
Call Cnv.To_dir_rodv
End Sub

Private Sub mnu_print_vk_Click()
Call Cnv.To_dir_vk
End Sub

Private Sub mnu_refresh_Click()
Call load_dirs
Call view_nonblock
End Sub
Private Sub check_dirs_p()
On Error Resume Next
For x = 1 To lstprnik.ListItems.Count
    Call mysql.query("select `idprnik` from `prnik_" & nowBase & "` where `fam` like '" & Left$(lstprnik.ListItems(x).SubItems(1), 4) & "%' and `txtvk` = '" & lstprnik.ListItems(x).SubItems(4) & "'")
    If val(st) > 0 Then
        lstprnik.ListItems(x).SubItems(6) = st
        lstprnik.ListItems(x).Bold = True
    End If
Next x
End Sub
Public Sub view_nonblock()
On Error Resume Next
Call mysql.query("select `fam`,`vk`,`did` from `directivi_p_" & nowBase & "` where `pid`='0'")
datt() = DAT()
stt = st
For x = 1 To stt
    Call mysql.query("select `idprnik` from `prnik_" & nowBase & "` where `fam` like '" & Left$(datt(1, x), 4) & "%' and `txtvk` = '" & datt(2, x) & "'")
    If val(st) > 0 Then
    For Y = 1 To 6
        lstdirs.FindItem(datt(3, x)).ListSubItems.Item(Y).ForeColor = &HFF&
    Next Y
    End If
Next x



End Sub

Private Sub search_Click()
frmdir_search.Show vbModal, Me
End Sub
