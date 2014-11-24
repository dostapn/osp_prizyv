VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   Caption         =   "Поиск"
   ClientHeight    =   11085
   ClientLeft      =   -2010
   ClientTop       =   -1800
   ClientWidth     =   14730
   ClipControls    =   0   'False
   DrawMode        =   16  'Merge Pen
   DrawStyle       =   5  'Transparent
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton kom 
      Caption         =   "Команда"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   2600
   End
   Begin VB.Frame Frame1 
      Caption         =   "    Дополнительные Параметры"
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   6360
      Width           =   2775
      Begin VB.OptionButton dreg 
         Caption         =   "Дата Регистрации"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   2400
      End
      Begin VB.OptionButton dvk 
         Caption         =   "Военкомат"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1320
         Width           =   2400
      End
      Begin VB.OptionButton dno 
         Caption         =   "Нет"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.OptionButton vb 
      Caption         =   "Военный Билет"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   2600
   End
   Begin VB.OptionButton rodv 
      Caption         =   "Род Войск"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   2600
   End
   Begin VB.OptionButton vch 
      Caption         =   "Воинская Часть"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   2600
   End
   Begin VB.OptionButton dataotp 
      Caption         =   "Дата Отправки"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   2600
   End
   Begin VB.OptionButton datapr 
      Caption         =   "Дата Призыва"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   2600
   End
   Begin VB.OptionButton vk 
      Caption         =   "Военкомат"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   2600
   End
   Begin VB.OptionButton fam 
      Caption         =   "Фамилия"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   2600
   End
   Begin VB.OptionButton ypk 
      Caption         =   "УПК"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   2600
   End
   Begin VB.CommandButton cmdOst 
      Caption         =   "Не в команды"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   9
      Top             =   8520
      Width           =   2600
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Удаленные"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   8
      Top             =   10320
      Width           =   2600
   End
   Begin VB.CommandButton cmdSochi 
      Caption         =   "СОЧИ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   7
      Top             =   9720
      Width           =   2600
   End
   Begin VB.CommandButton cmdVozvrat 
      Caption         =   "Возврат"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   6
      Top             =   9120
      Width           =   2600
   End
   Begin VB.TextBox txtdop 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   2600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":6852
            Key             =   "ADM_LOCK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":712C
            Key             =   "LOCK"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":7A06
            Key             =   "LOCK_COMM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":86E0
            Key             =   "OK"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtpole 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Left            =   240
      TabIndex        =   0
      Text            =   "                                                                                                                  "
      Top             =   240
      Width           =   2600
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   10815
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView listres 
      Height          =   10815
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   19076
      SortKey         =   5
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483641
      BackColor       =   12648384
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "УПК"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Фамилия"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Имя"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Отчество"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Военкомат"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Дата привоза"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Директва"
         Object.Width           =   265
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H80000018&
      Caption         =   "Найти"
      Default         =   -1  'True
      DisabledPicture =   "frmSearch.frx":8FBA
      DownPicture     =   "frmSearch.frx":F80C
      Height          =   345
      Left            =   4440
      Picture         =   "frmSearch.frx":1605E
      TabIndex        =   3
      Top             =   3360
      Width           =   1755
   End
   Begin VB.ComboBox lstvk 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2600
   End
   Begin VB.Label info_t 
      Caption         =   "Label1"
      Height          =   15
      Left            =   360
      TabIndex        =   23
      Top             =   5160
      Width           =   15
   End
   Begin VB.Menu nmuResoult 
      Caption         =   "Результат"
      Begin VB.Menu mnuToExcel 
         Caption         =   "В Excel"
         Shortcut        =   +{F12}
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDel_Click()
Call clr
Call ostsearch
listres.ColumnHeaders.Item(5) = "Военкомат"
listres.ColumnHeaders.Item(6) = "Дата Удаления"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,dataosp FROM delprnik_" & nowBase & " WHERE idprnik like"
    info_t.Caption = "1"
    table = "delprnik"
    Call cmdSearch_Click
End Sub
Private Sub ostsearch()
ypk = False
fam = False
vk = False
datapr = False
dataotp = False
vch = False
rodv = False
vb = False
dno = True
End Sub
Private Sub cmdOst_Click()
Call clr
Call ostsearch
listres.ColumnHeaders.Item(5) = "Военкомат"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,dataosp FROM prnik_" & nowBase & " WHERE otprvid ='0'"
    info_t.Caption = "0"
    Call cmdSearch_Click
End Sub

Public Sub cmdSearch_Click()
On Error Resume Next
frmSearch.Caption = "Поиск: Обработка запроса"

listres.ListItems.Clear
append = appendSQL()
If datapr = True Or dataotp = True Then txtpole = CnvDataWinToSql(txtpole)
Call mysql.query(sql_com & " '" & txtpole & "%' " & append & " ORDER by txtvk,fam,name,otch")
datt() = DAT()


    For x = 1 To st
        lock_tp = datt(1, x)
        If lock_tp = "0" Then Set LF = listres.ListItems.add(, , datt(2, x), "OK", "OK")
        If lock_tp = "1" Then Set LF = listres.ListItems.add(, , datt(2, x), "LOCK", "LOCK")
        If lock_tp = "2" Or lock_tp = "4" Then Set LF = listres.ListItems.add(, , datt(2, x), "LOCK_COMM", "LOCK_COMM")
        If lock_tp = "3" Then Set LF = listres.ListItems.add(, , datt(2, x), "ADM_LOCK", "ADM_LOCK")
           For Y = 1 To 5
                If IsDate(datt(Y + 2, x)) Then datt(Y + 2, x) = CnvDataSqLToWin(datt(Y + 2, x))
                LF.SubItems(Y) = datt(Y + 2, x)
            Next Y
            'LF.ToolTipText = get_info_prnik(datt(2, x))
    Next x
         If datapr = True Or dataotp = True Then txtpole = CnvDataSqLToWin(txtpole)
         SB1.Panels(1) = "Всего:"
         SB1.Panels(2) = listres.ListItems.Count
         

Call ReSizeColumnHeaders(listres)
frmSearch.Caption = "Поиск: результат в " & listres.ListItems.Count & " призывников."
txtpole.SelStart = 0: txtpole.SelLength = Len(txtpole)

End Sub

Private Sub cmdSochi_Click()
Call clr
Call ostsearch
listres.ColumnHeaders.Item(5) = "Военкомат"
listres.ColumnHeaders.Item(6) = "Сбежал"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,datedel FROM sochi_" & nowBase & " WHERE idprnik like"
         datetypegen = 7
         info_t.Caption = "1"
         table = "sochi"
    Call cmdSearch_Click
End Sub

Private Sub cmdVozvrat_Click()
Call clr
Call ostsearch
listres.ColumnHeaders.Item(5) = "Военкомат"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,datedel FROM vozvrat_" & nowBase & " WHERE idprnik like"
    table = "vozvrat"
    info_t.Caption = "1"
    datetypegen = 7
Call cmdSearch_Click
End Sub

Private Sub dataotp_Click()
Call clr
txtpole = Date
listres.ColumnHeaders.Item(5) = "Военкомат"
listres.ColumnHeaders.Item(5) = ""
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk, FROM prnik_" & nowBase & ",otpravka_" & nowBase & " WHERE prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid AND data like "
    info_t.Caption = "0"
    datetypegen = 7
    txtpole.SetFocus
End Sub

Private Sub datapr_Click()
Call clr
txtpole = Date
listres.ColumnHeaders.Item(5) = "Военкомат"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,dataosp FROM prnik_" & nowBase & " WHERE dataosp ="
      info_t.Caption = "0"
      datetypegen = 7
      txtpole.SetFocus
      
End Sub

Private Sub dreg_Click()
On Error Resume Next
txtdop = Date
End Sub

Private Sub fam_Click()
On Error Resume Next
Call clr
 listres.ColumnHeaders.Item(5) = "Военкомат"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,dataosp FROM prnik_" & nowBase & " WHERE fam like"
    info_t.Caption = "0"
    datetypegen = 7

 If errf = False Then txtpole.SetFocus
 errf = False
End Sub

Private Sub Form_Load()

On Error Resume Next
Call sorting_blok("txtvk", "prnik_" & nowBase)
    For x = 1 To UBound(outm())
        lstvk.AddItem (outm(x))
        
    Next x

dno = True
errf = True
fam = True

End Sub
Private Sub Form_Resize()
On Error Resume Next
    listres.Move 3120, 0, Me.ScaleWidth - 3200, Me.Height - 1000
    Call ReSizeColumnHeaders(listres)
End Sub

Function appendSQL() As String
    If dvk = True Then appendSQL = " AND txtvk like '" & txtdop & "%'"
    If dreg = True Then appendSQL = " AND dataosp like '" & CnvDataWinToSql(txtdop) & "%'"
End Function

Private Sub Form_Unload(Cancel As Integer)

'Call mysql.free_result
Set frmSearch = Nothing
End Sub

Private Sub kom_Click()
Call clr
sql_com = "SELECT `lock`,idprnik,fam, name, otch, txtvk, oblkom, vch FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid AND prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND naryad_" & nowBase & ".oblkom like"
listres.ColumnHeaders.Item(6) = "Род Войск"
txtpole.SetFocus
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

Private Sub listres_DblClick()
On Error Resume Next
If acl = "G" Or acl = "D" Or acl = "O" Then
If info_t.Caption = "0" Then
    expupk = VAl(listres.ListItems(listres.SelectedItem.Index).Text)
    frmInfoPr.Show vbModal, Me
End If
If info_t.Caption = "1" Then
    expupk = VAl(listres.ListItems(listres.SelectedItem.Index).Text)
    frmdelinfo.Show vbModal, Me
End If
End If
End Sub

Private Sub listres_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

Private Sub listres_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 4 Then
Dim str As Long
str = InputBox("asd", "asd")

For x = listres.ListItems.Count To 1 Step -1
              If listres.ListItems(x).Selected Then Call mysql.query("UPDATE prnik_" & nowBase & " set otprvid ='" & str & "' WHERE idprnik='" & listres.ListItems(x).Text & "'")
Next x
End If
End Sub

Private Sub lstvk_Click()
txtpole.Text = lstvk.Text

End Sub

Private Sub mnuToExcel_Click()
Call Cnv.ResoultSearch(listres, True, "Список", Caption)
End Sub

Private Sub rodv_Click()
Call clr
sql_com = "SELECT `lock`,idprnik,fam, name, otch, txtvk, rodv, vch FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid AND prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND naryad_" & nowBase & ".rodv like"
datetypegen = 0
info_t.Caption = "0"
listres.ColumnHeaders.Item(6) = "Род Войск"
txtpole.SetFocus
End Sub

Private Sub txtpole_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub
Private Sub vib_dop_Click()
    Select Case vib_dop.ListIndex
        Case 1
        txtdop.Visible = False
        Case Else
        txtdop.Visible = True
    End Select
End Sub

Private Sub vb_Click()
Call clr
txtpole.SetFocus
End Sub

Private Sub vch_Click()
Call clr
sql_com = "SELECT `lock`,idprnik, fam,name,otch,txtvk,naryad_" & nowBase & ".vch FROM prnik_" & nowBase & ", otpravka_" & nowBase & ", naryad_" & nowBase & " WHERE naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid AND prnik_" & nowBase & ".otprvid = otpravka_" & nowBase & ".otpravkaid AND naryad_" & nowBase & ".vch like"
info_t.Caption = "0"
datetypegen = 7
txtpole.SetFocus
End Sub


Private Sub vk_Click()
txtpole = vbNullString
lstvk.Visible = True
txtpole.Visible = False
listres.ColumnHeaders.Item(5) = "Год Рождения"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,datar,dataosp FROM prnik_" & nowBase & " WHERE txtvk like "
  info_t.Caption = "0"
  datetypegen = 7
End Sub
Private Sub clr()
On Error Resume Next
txtpole = vbNullString
lstvk.Visible = False
txtpole.Visible = True
'
txtpole.SelStart = 0: txtpole.SelLength = Len(txtpole)
End Sub
Private Sub ypk_Click()
Call clr
listres.ColumnHeaders.Item(5) = "Военкомат"
    sql_com = "SELECT `lock`,idprnik,fam,name,otch,txtvk,dataosp FROM prnik_" & nowBase & " WHERE idprnik like "
info_t.Caption = "0"
datetypegen = 7
txtpole.SetFocus
End Sub
