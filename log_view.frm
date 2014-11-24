VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLog 
   Caption         =   "Отчет"
   ClientHeight    =   12150
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   13830
   Icon            =   "log_view.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12150
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   480
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker to_date 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49020929
      CurrentDate     =   39253
   End
   Begin MSComCtl2.DTPicker from_date 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49020929
      CurrentDate     =   39253
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Сгенерировать отчет"
      Height          =   735
      Left            =   11880
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ListView listres 
      Height          =   10095
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   17806
      SortKey         =   1
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Дата"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Время"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Ф.И.О."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Действие к"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Действие"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Д"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "2"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "по"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "c"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommAND1_Click()
On Error Resume Next
Dim datt() As String
Dim in_id As Long
Call log_types
listres.ListItems.Clear
Call mysql.query("SELECT `data`,`time`,`who`,`type`,`act`,`argw`,`to_id` FROM logs_" & nowBase & " WHERE data >= '" & CnvDataWinToSql(from_date) & "' AND data <= '" & CnvDataWinToSql(to_date) & "'")
datt() = DAT()
For x = 1 To st
    datt(1, x) = CnvDataSqLToWin(datt(1, x))
    datt(3, x) = get_fio(datt(3, x))
    in_id = datt(7, x)
    If datt(4, x) = "0" Then datt(7, x) = log_get_fio(in_id)
    If datt(4, x) = "1" Then datt(7, x) = log_get_kom(in_id)
    'If datt(4, x) = "2" Then datt(7, x) = log_get_nar(in_id)
    datt(4, x) = log_type_act(datt(4, x))
    datt(5, x) = log_act(datt(5, x))
   
    
    Set LF = listres.ListItems.add()
    For c = 1 To 7
    LF.SubItems(c) = datt(c, x)
    
    Next c
Next x
Call ReSizeColumnHeaders(listres)
listres.Refresh
End Sub

Private Sub Form_Load()
from_date = Date
to_date = Date
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
