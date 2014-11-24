VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmdir_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Поиск директивщиков"
   ClientHeight    =   5010
   ClientLeft      =   6015
   ClientTop       =   6075
   ClientWidth     =   8535
   Icon            =   "frmdir_search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Поиск"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ф.И.О"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "г.р."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Военкомат"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Директива"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtFam 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Фамилия"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmdir_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
Call mysql.query("select directivi_" & nowBase & ".`id`,`fam`,`name`,`otch`,`year_b`,`vk`,`dir`,`did` from `directivi_" & nowBase & "`,`directivi_p_" & nowBase & "` where `fam` like '" & txtFam & "%' and directivi_p_" & nowBase & ".`did`=directivi_" & nowBase & ".`id`")
For x = 1 To st
    Set LF = lstSearch.ListItems.add(, , DAT(1, x))
    LF.SubItems(1) = DAT(2, x) & " " & Left(DAT(3, x), 1) & ". " & Left(DAT(4, x), 1) & "."
    LF.SubItems(2) = DAT(5, x)
    LF.SubItems(3) = DAT(6, x)
    LF.SubItems(4) = DAT(7, x)
    LF.SubItems(5) = DAT(8, x)
    
Next x
End Sub

Private Sub lstSearch_DblClick()
frmdir.lstdirs.FindItem(lstSearch.SelectedItem.SubItems(5)).Selected = True
frmdir.lstdirs_Click
frmdir.lstdirs.FindItem(lstSearch.SelectedItem.SubItems(5)).EnsureVisible
frmdir.lstprnik.FindItem(lstSearch.SelectedItem.Text).Selected = True
frmdir.lstprnik.FindItem(lstSearch.SelectedItem.Text).EnsureVisible
Unload Me
End Sub
