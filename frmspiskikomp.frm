VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmspiskikomp 
   Caption         =   "����� �� ������������������"
   ClientHeight    =   9075
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11730
   Icon            =   "frmspiskikomp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   10320
      TabIndex        =   18
      Top             =   6720
      Width           =   1335
      Begin VB.OptionButton opt2_2 
         Caption         =   "�������"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton opt2_1 
         Caption         =   "���"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "����������"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   9000
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
      Begin VB.OptionButton opt1_3 
         Caption         =   "����."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton opt1_2 
         Caption         =   "�� ����."
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton opt1_1 
         Caption         =   "����"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "����������"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����������"
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   8895
      Begin VB.CommandButton cmdinfo 
         Caption         =   "����������"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdblock 
         Caption         =   "�������������"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox ypk_vibor 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmspiskikomp.frx":6852
         Left            =   240
         List            =   "frmspiskikomp.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "���� ��������:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "���� �������:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "�.�.�.:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "��:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label privoz 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label txtvk 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label gr 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label txtname 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2640
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "���"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView listres 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�������, ���, ��������"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�.�."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "#"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "B"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.Menu mnu_add_pr 
      Caption         =   "��������"
   End
   Begin VB.Menu mnu_rk 
      Caption         =   "���"
      Visible         =   0   'False
      Begin VB.Menu red 
         Caption         =   "�������������"
      End
      Begin VB.Menu del_pr 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "frmspiskikomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_com As String
Dim whm(3) As String
Dim type_gen, wh As String

Private Sub cmdBlock_Click()
Call mysql.query("update `prnik_" & nowBase & "` set `lock`='1', `lprim`='" & listres.SelectedItem.SubItems(6) & " " & listres.SelectedItem.SubItems(5) & "' where `idprnik`='" & ypk_vibor.Text & "'")
MsgBox "��������� " & listres.SelectedItem.SubItems(2) & " ������������.", OKOnly
Call mysql.query("update spiski_" & nowBase & " set `lock_id`='" & ypk_vibor.Text & "', `lock_pr` = 1 where `id`='" & listres.ListItems(listres.SelectedItem.Index) & "'")
Call Form_Load
End Sub

Private Sub cmdinfo_Click()
On Error Resume Next
expupk = ypk_vibor.Text
If expupk > 0 Then frmInfoPr.Show vbModal, Me
End Sub

Private Sub del_pr_Click()
If acl = "G" Or acl = "O" Then
If MsgBox("������������� �������� ���������� " & listres.SelectedItem.SubItems(2) & "?", vbYesNo + vbQuestion, strMAIN_TITLE) = vbYes Then
mysql.query ("delete from spiski_" & nowBase & " where spiski_" & nowBase & ".id = " & listres.SelectedItem.SubItems(1) & "")
End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call whm_in
Call opt1_1_Click
Call opt2_1_Click
Call gen_list(type_gen, wh)
End Sub
Public Sub gen_list(type_gen, wh As String)
On Error Resume Next
listres.ListItems.Clear
If type_gen = "" Then Call mysql.query("Select `id`,`fio`,`gr`,`vk`,`kyda`,`kto`,`type_b`,`lock_pr`,`lock_id` from spiski_" & nowBase & "  ORDER BY `fio`,`vk`,`gr`")
If type_gen = "0" Then Call mysql.query("Select `id`,`fio`,`gr`,`vk`,`kyda`,`kto`,`type_b`,`lock_pr`,`lock_id` from spiski_" & nowBase & " where lock_id = '0' ORDER BY `fio`,`vk`,`gr`")
If type_gen = "1" Then Call mysql.query("Select `id`,`fio`,`gr`,`vk`,`kyda`,`kto`,`type_b`,`lock_pr`,`lock_id` from spiski_" & nowBase & " where lock_id > '0' ORDER BY `fio`,`vk`,`gr`")
datt() = DAT()
stt = st
Dim dataosp_1 As String
dataosp_1 = Format(Now, "YYYY-MM-DD")
If stt > 0 Then
    For x = 1 To stt
        Set LF = listres.ListItems.add(, , datt(1, x))
        For Y = 1 To 6
               LF.SubItems(Y) = datt(Y, x)
             
               
        Next Y
        LF.SubItems(7) = whm(datt(7, x))
If wh = "" Then Call mysql.query("select `idprnik` from prnik_" & nowBase & " where fam like '" & Left(datt(2, x), 4) & "%'")
If wh = "0" Then Call mysql.query("select `idprnik` from prnik_" & nowBase & " where fam like '" & Left(datt(2, x), 4) & "%' and dataosp = '" & dataosp_1 & "'")
        If st > 0 Then
        LF.SubItems(8) = st
        listres.ListItems(x).ListSubItems(1).ForeColor = 255
        listres.ListItems(x).ListSubItems(2).ForeColor = 255
        listres.ListItems(x).ListSubItems(3).ForeColor = 255
        listres.ListItems(x).ListSubItems(4).ForeColor = 255
        listres.ListItems(x).ListSubItems(5).ForeColor = 255
        listres.ListItems(x).ListSubItems(6).ForeColor = 255
        
        End If
        If datt(9, x) > 0 Then LF.SubItems(9) = "��"
        
    Next x

End If

End Sub
Private Sub listres_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    If Button = 2 Then Call PopupMenu(mnu_rk)
    
End Sub
Private Sub listres_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
        If KeyCode = 40 Then Call listres_Click
        listres.SelectedItem.Index() = (listres.SelectedItem.Index() + 1)
        If KeyCode = 38 Then Call listres_Click
End Sub
Public Sub gen_list_all(type_gen, wh As String)
On Error Resume Next
listres.ListItems.Clear
Call mysql.query("Select `id`,`fio`,`gr`,`vk`,`kyda`,`kto`,`type_b`,`lock_pr`,`lock_id` from spiski_" & nowBase & " where lock_pr like '" & type_gen & "%' and `type_b` like '" & wh & "%' ORDER BY `fio`,`vk`,`gr`")
datt() = DAT()
stt = st
If stt > 0 Then
    For x = 1 To stt
        Set LF = listres.ListItems.add(, , datt(1, x))
        For Y = 1 To 6
               LF.SubItems(Y) = datt(Y, x)
        Next Y
        LF.SubItems(7) = whm(datt(7, x))
        Call mysql.query("select `idprnik` from prnik_" & nowBase & " where fam like '" & Left(datt(2, x), 4) & "%'")
        If st > 0 Then
        LF.SubItems(8) = st
        LF.ForeColor = 255
        'LF.ForeColor = 255
        listres.ListItems(Index).ForeColor = 255
        End If
        If datt(9, x) > 0 Then LF.SubItems(9) = "��"
    Next x

End If
End Sub
Private Sub whm_in()
whm(0) = "���"
whm(1) = "���"
whm(2) = "���"
whm(3) = "��������"
End Sub

Private Sub Form_Resize()
listres.Move 0, 0, ScaleWidth, Me.Height - 3100
frame1.Move 0, Me.Height - 3000
Frame2.Move 9000, Me.Height - 3000
Frame3.Move 10320, Me.Height - 3000
Call ReSizeColumnHeaders(listres)
End Sub

Private Sub listres_Click()
On Error Resume Next
Dim dataosp_1 As String
dataosp_1 = Format(Now, "YYYY-MM-DD")
ypk_vibor.Clear
If listres.SelectedItem.SubItems(8) > 0 Then
If wh = "" Then Call mysql.query("select `idprnik` from prnik_" & nowBase & " where fam like '" & Left(listres.SelectedItem.SubItems(2), 4) & "%'")
    If wh = "0" Then Call mysql.query("select `idprnik` from prnik_" & nowBase & " where fam like '" & Left(listres.SelectedItem.SubItems(2), 4) & "%' and dataosp = '" & dataosp_1 & "'")
        ypk_vibor.Enabled = True
             For x = 1 To st
                ypk_vibor.AddItem (DAT(1, x))
                
                Next x
                ypk_vibor.ListIndex = 0
        Call ypk_vibor_Click
End If
End Sub




Private Sub opt1_1_Click()
type_gen = ""
Call gen_list(type_gen, wh)
End Sub

Private Sub opt1_2_Click()
type_gen = "0"
Call gen_list(type_gen, wh)
End Sub

Private Sub opt1_3_Click()
type_gen = "1"
Call gen_list(type_gen, wh)
End Sub

Private Sub opt2_1_Click()
wh = ""
Call gen_list(type_gen, wh)
End Sub

Private Sub opt2_2_Click()
wh = "0"
Call gen_list(type_gen, wh)
End Sub



Private Sub red_Click()
If acl = "G" Or acl = "O" Then
frmspiski_add.Add_pr.Caption = "���������"
frmspiski_add.txtid.Caption = listres.SelectedItem.SubItems(1)
frmspiski_add.txtfio.Text = listres.SelectedItem.SubItems(2)
frmspiski_add.Years.Text = listres.SelectedItem.SubItems(3)
frmspiski_add.lstvk.Text = listres.SelectedItem.SubItems(4)
frmspiski_add.txtto.Text = listres.SelectedItem.SubItems(5)
frmspiski_add.txtkto.Text = listres.SelectedItem.SubItems(6)

frmspiski_add.Show vbModal, Me
End If
End Sub

Private Sub ypk_vibor_Click()
On Error Resume Next
Call mysql.query("select concat(fam,' ', name,' ', otch), datar, txtvk, dataosp from prnik_" & nowBase & " where idprnik = '" & ypk_vibor.Text & "'")
txtname.Caption = DAT(1, 1)
gr.Caption = DAT(2, 1)
txtvk.Caption = DAT(3, 1)
privoz.Caption = DAT(4, 1)
End Sub

Private Sub mnu_add_pr_Click()
If acl = "G" Or acl = "O" Then
frmspiski_add.Show vbModal, Me
End If
End Sub
