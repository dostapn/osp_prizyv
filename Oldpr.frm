VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Oldpr 
   Caption         =   "Form2"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13635
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Искать"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox search 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin MSComctlLib.ListView listres 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   16748
      SortKey         =   2
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Фамилия"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Имя"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Отчество"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Год Рождения"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Военкомат"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Пункт Дислокации"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Род войск"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "В/Ч"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Отправка"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Width           =   4
      EndProperty
   End
End
Attribute VB_Name = "Oldpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim x As Long

'ListFiles.ListItems.Clear
listres.ListItems.Clear
Call mysql.query("select * from olddb where fam like '" & search.Text & "%'")

For x = 1 To st
    Set LF = listres.ListItems.Add()
    For c = 1 To 10
   LF.SubItems(c) = DAT(c + 1, x)
    Next c
Next x
listres.Refresh
End Sub

Private Sub Form_Load()
Set mysql = New cMysql
z.txtDB = "prizyv_olddb"
z.txtUser = "root"
z.txtPass = "not-4-all"
mysql.real_connect z.txtHost, z.txtUser, z.txtPass, z.txtDB, CLng(Val(z.txtPort)), , 0
Oldpr.Caption = "База: " & z.txtDB & " Пользователь: " & lgn
Oldpr.Refresh

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
