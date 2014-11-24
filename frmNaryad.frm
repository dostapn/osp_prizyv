VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmNaryad 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Наряд"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   -900
   ClientWidth     =   13065
   Icon            =   "frmNaryad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbalListViewLib6.vbalListViewCtl listres 
      Height          =   10455
      Left            =   1920
      TabIndex        =   11
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   18441
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   4
      LabelEdit       =   0   'False
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
      DoubleBuffer    =   -1  'True
      Begin vbalIml6.vbalImageList img1 
         Left            =   1320
         Top             =   3960
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   24
         IconSizeY       =   24
         ColourDepth     =   24
         Size            =   2460
         Images          =   "frmNaryad.frx":6852
         Version         =   131072
         KeyCount        =   1
         Keys            =   ""
      End
   End
   Begin ComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   10500
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   10
   End
   Begin VB.Frame Frame1 
      Caption         =   "Быстрый поиск"
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
      Begin VB.TextBox txtsearch_data 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtsearch_okrkom 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtsearch_pred 
         BackColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtsearch_punkt 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtsearch_vch 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtsearch_oblkom 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "По дате"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   4560
         Width           =   615
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label5 
         Caption         =   "По ОКР. команде"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "По предназначению"
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "По ДИСЛОКАЦИИ"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "По ВЧ"
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
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   1800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "По ОБЛ. команде"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.TreeView rodvtr 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   8705
      _Version        =   393217
      LabelEdit       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mnu_rodv 
      Caption         =   "Наряд"
      Begin VB.Menu mnu_com_add 
         Caption         =   "Добавить команду"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_update 
         Caption         =   "Обновить"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnu_print 
      Caption         =   "Печать"
      Begin VB.Menu mnu_print_select_naryad 
         Caption         =   "Выбранный наряд"
      End
      Begin VB.Menu mnu_print_res_search 
         Caption         =   "Результата поиска"
      End
      Begin VB.Menu mnu_print_line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print_srezki 
         Caption         =   "Срезки"
      End
      Begin VB.Menu mnu_print_dolgi 
         Caption         =   "Долги"
      End
      Begin VB.Menu mnu_print_ostatki 
         Caption         =   "Остатки"
      End
      Begin VB.Menu mnu_print_line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print_podwef 
         Caption         =   "Подшефные части"
      End
      Begin VB.Menu mnu_print_zato 
         Caption         =   "ЗАТО"
      End
      Begin VB.Menu mnu_print_cnp 
         Caption         =   "ЦНП"
      End
      Begin VB.Menu mnu_print_prim 
         Caption         =   "Примечания"
      End
      Begin VB.Menu mnu_print_line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print_nar_full 
         Caption         =   "Наряд полный"
      End
      Begin VB.Menu mnu_print_nar_days 
         Caption         =   "Наряд по дням"
      End
      Begin VB.Menu mnu_print_nar_date 
         Caption         =   "Наряд на дату"
      End
   End
   Begin VB.Menu mnu_service 
      Caption         =   "Сервис"
      Begin VB.Menu mnu_rebind_oblkom 
         Caption         =   "Переназначение областных команд"
      End
      Begin VB.Menu mnu_block_nar 
         Caption         =   "Заблокировать наряд"
      End
   End
   Begin VB.Menu menu_del_kom 
      Caption         =   "menu_del"
      Visible         =   0   'False
      Begin VB.Menu mnu_add_komm_menu 
         Caption         =   "Добавить команду"
      End
      Begin VB.Menu mnu_edit_komm_menu 
         Caption         =   "Редактировать/Показать"
      End
      Begin VB.Menu mnu_del_kom 
         Caption         =   "Удалить команду"
      End
      Begin VB.Menu add_kom_now_day 
         Caption         =   "В команды на сегодня"
      End
   End
End
Attribute VB_Name = "frmNaryad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itmX As cListItem
Dim sitmX As cListItems
Dim sFnt As New StdFont
Dim colX As cColumn
Private Sub Form_Load()

Call load_listress
If get_VAlue("block_nar") = "1" Then mnu_rebind_oblkom.Enabled = False: mnu_block_nar.Caption = "Разблокировать Наряд(ОСТОРОЖНО)"
If get_VAlue("block_nar") = "0" Then mnu_rebind_oblkom.Enabled = True: mnu_block_nar.Caption = "Заблокировать Наряд"
Call rodv_refresh
Call bar_refresh
txtsearch_data = CnvDataWinToSql(Date)
End Sub
Public Sub bar_refresh()
On Error Resume Next
Call mysql.query("select sum(kolvo) from naryad_" & nowBase)
stbar.Panels(1) = " Всего наряд : " & dat(1, 1)
x = dat(1, 1)
Call mysql.query("select sum(kolvo) from naryad_srezki_" & nowBase)
stbar.Panels(2) = " Отправлено всего : " & dat(1, 1)
Y = dat(1, 1)
stbar.Panels(3) = " Осталось отправить : " & x - Y
Dim lp() As String
lp = Split(procent_otp * 100, ",")
stbar.Panels(4) = " В процентах : " & lp(0) & "," & Left(lp(1), 2) & "%"
Call mysql.query("select sum(dolg) from naryad_" & nowBase)
stbar.Panels(5) = " Общий долг : " & dat(1, 1)
Call mysql.query("select sum(kolvo) from naryad_" & nowBase & " where `datanar`='" & CnvDataWinToSql(Date) & "'")
If VAl(dat(1, 1)) = 0 Then dat(1, 1) = "0"
stbar.Panels(6) = " Сегодня по плану : " & dat(1, 1)
End Sub

Private Sub listres_ItemDblClick(Item As vbalListViewLib6.cListItem)
p_com_id = VAl(listres.ListItems.Item(listres.SelectedItem.Index).SubItems(1).Caption)
frmnaryad_info.Show vbModal, Me
End Sub

Private Sub mnu_print_cnp_Click()
Call Cnv.To_naryad_rz("4")
End Sub

Private Sub mnu_print_prim_Click()
Call Cnv.To_prim
End Sub

Private Sub mnu_rebind_oblkom_Click()
''''''''''''''' переназначение областных команд согласно narid
If acl = "G" Then
Call mysql.query("update naryad_" & nowBase & " set oblkom = narid")
End If
End Sub

Private Sub stbar_click()
Call bar_refresh
End Sub
Public Sub rodv_refresh()

On Error Resume Next
Dim nodX As Node
Dim datt() As String
Dim stt As Long
rodvtr.Nodes.Clear
  Set nodX = rodvtr.Nodes.add(, , "r0", "Наряд целеком")
 Call mysql.query("select `okr` from naryad_" & nowBase & " group by `okr`")
 datt() = dat()
 stt = 1
    For x = 1 To st
      Set nodX = rodvtr.Nodes.add(, , "r" & x, datt(1, x))
        Call mysql.query("SELECT `rodv` FROM naryad_" & nowBase & " WHERE okr='" & datt(1, x) & "' group by rodv")
               If st > 0 Then
               
                For Y = 1 To st
                
                    Set nodX = rodvtr.Nodes.add("r" & x, tvwChild, "c" & stt, dat(1, Y))
                    nodX.EnsureVisible
                    stt = stt + 1
                Next Y
            End If
            nodX.EnsureVisible
    Next x
  nodX.EnsureVisible
rodvtr.Nodes(1).Selected = True
listres.View = eViewDetails
      listres.CustomDraw = True
      listres.AutoArrange = True
      listres.OneClickActivate = True
'Call rodvtr_Click
End Sub
Private Sub Form_Resize()
On Error Resume Next
listres.Move 1920, 0, Me.Width - 2000, Me.Height - 1100
Call ReSizeColumnHeaders_new(listres)
End Sub

Private Sub listres_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Unload Me
End Sub

Private Sub listres_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 2 Then Call PopupMenu(menu_del_kom)
End Sub

Private Sub mnu_add_komm_menu_Click()
Call mnu_com_add_Click
End Sub

Private Sub mnu_block_nar_Click()
If acl = "G" Then
If get_config("block_nar") = "0" Then
    If MsgBox("Вы уверены что хотите заблокировать наряд?" & Chr(10) & "После блокировки вы не сможете переназначать номера областных команд", vbInformation + vbYesNo) = vbYes Then
        Call mysql.query("UPDATE bases set `block_nar`='1' WHERE `VAl`='" & nowBase & "'")
        mnu_rebind_oblkom.Enabled = False
    End If
Else
If MsgBox("Вы уверены что хотите разблокировать наряд?" & Chr(10) & "После разблокировки вы можете повредить таблицы нарядов, что приведет к искревлению и неправильной работе все системы!!!" & Chr(10) & "!!!НИКОГДА не разблокируйте наряд во время ПРИЗЫВА и ОТПРАВКИ В ВОЙСКА!!!", vbInformation + vbYesNo) = vbYes Then
        Call mysql.query("UPDATE bases set `block_nar`='0' WHERE `VAl`='" & nowBase & "'")
        mnu_rebind_oblkom.Enabled = True
    End If

End If
End If
End Sub
Private Sub mnu_com_add_Click()
If acl = "G" Then
On Error Resume Next
frmnaryad_info.Caption = "Добавить команду"
frmnaryad_info.cmd_save.Caption = "Добавить"

If rodvtr.SelectedItem.Parent = "nothing" Then
Else
frmnaryad_info.txtokr = rodvtr.SelectedItem.Parent
frmnaryad_info.txtpodrod = rodvtr.SelectedItem.Text
End If
frmnaryad_info.txtoblkom.Enabled = True
frmnaryad_info.Show vbModal, Me
End If
End Sub

Private Sub mnu_del_kom_Click()
On Error Resume Next
If MsgBox("Вы уверены что хотите удалить команду " & listres.ListItems.Item(listres.SelectedItem.Index).SubItems(3).Caption & "?", vbQuestion + vbYesNo) = vbYes Then
    Call mysql.query("DELETE FROM naryad_" & nowBase & " WHERE narid='" & listres.ListItems(listres.SelectedItem.Index).SubItems(1).Caption & "'")
    Call mysql.query("DELETE FROM naryad_srezki_" & nowBase & " WHERE `narid`='" & listres.ListItems(listres.SelectedItem.Index).SubItems(1).Caption & "'")
    Call rodvtr_Click
End If
End Sub

Private Sub mnu_edit_komm_menu_Click()
p_com_id = VAl(listres.ListItems.Item(listres.SelectedItem.Index).SubItems(1).Caption)
frmnaryad_info.Show vbModal, Me
End Sub

Private Sub mnu_print_dolgi_Click()
Call Cnv.To_naryad_rz("1")
End Sub
Private Sub add_kom_now_day_Click()


Dim indSel As Long
            CLICK_COM = listres.ListItems(listres.SelectedItem.Index).SubItems(1).Caption
            Dim new_id As Long
            Dim NEW_NAR_ID As Long
            dateotp = frmMain.cal.Year & "-" & Format(frmMain.cal.Month, "00") & "-" & Format(frmMain.cal.Day, "00")
            If Not mysql.query("SELECT max(otpravkaid) FROM otpravka_" & nowBase & "") Then
                MsgBox "Connect ErrOR", vbCritical, strMAIN_TITLE
                Exit Sub
            End If
            
            If Not st = 0 Then new_id = dat(1, st) + 1 Else new_id = 1
            
            Call mysql.query("SELECT narid, oblkom FROM naryad_" & nowBase & " WHERE oblkom = '" & CLICK_COM & "'")
            For x = 1 To st
            If InStr(1, dat(2, x), CLICK_COM, vbBinaryCompare) > 0 Then NEW_NAR_ID = CLng(dat(1, x)): GoTo Cont
            Next x
            MsgBox "Ooops!", vbCritical, "STOP": Exit Sub
            
Cont:
            If new_id = 0 Then new_id = 1
            Call mysql.query("INSERT INTO `otpravka_" & nowBase & "` ( `otpravkaid` , `data` , `narid` , `fORpunkt` , `fORchast` , `kolvo` )VALUES ('" & new_id & "','" & dateotp & "','" & NEW_NAR_ID & "','" & txtForP & "','" & txtForCh & "','0');")
            
            Call log_sql("1", "7", CLICK_COM, "")
            
MsgBox "Команда №" & listres.ListItems.Item(listres.SelectedItem.Index).SubItems(3).Caption & " добавлена.", OKOnly

        
        
        
    End Sub

Private Sub mnu_print_nar_date_Click()
Call Cnv.To_naryad_date_d
End Sub

Private Sub mnu_print_nar_days_Click()
Call Cnv.To_naryad_date
End Sub

Private Sub mnu_print_nar_full_Click()
Call Cnv.To_naryad_all("")
End Sub

Private Sub mnu_print_ostatki_Click()
Call Cnv.To_naryad_ostatki
End Sub

Private Sub mnu_print_podwef_Click()
Call Cnv.To_naryad_rz("2")
End Sub

Private Sub mnu_print_res_search_Click()
Call Cnv.TO_naryad_search
End Sub

Private Sub mnu_print_SELECT_naryad_Click()

If rodvtr.SelectedItem.Index = "1" Then
Call Cnv.To_naryad_all("")
Else
If rodvtr.SelectedItem.Expanded = False And rodvtr.SelectedItem.Index > 1 Then
Call Cnv.To_naryad_all("and rodv = '" & rodvtr.SelectedItem & "' and okr = '" & rodvtr.SelectedItem.Parent & "'")
End If
If rodvtr.SelectedItem.Index > 1 And rodvtr.SelectedItem.Parent = "nothing" Then
Call Cnv.To_naryad_all("and okr = '" & rodvtr.SelectedItem & "'")
End If
End If

End Sub

Private Sub mnu_print_srezki_Click()
frmnaryad_srezki_sel.Show vbModal, Me
End Sub

Private Sub mnu_print_zato_Click()
Call Cnv.To_naryad_rz("3")
End Sub

Private Sub mnu_UPDATE_Click()
Call rodv_refresh
End Sub

Private Sub load_listress()


With listres
      .GridLines = True
      .FullRowSelect = True
      .HideSelection = False
      
      Set colX = .Columns.add(, , "narid")
      colX.Width = "1"
      Set colX = .Columns.add(, , "narid")
      colX.Width = "1"
      Set colX = .Columns.add(, , "ОМУ")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "ОСП")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "800"
      Set colX = .Columns.add(, , "Род")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "800"
      Set colX = .Columns.add(, , "Округ")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "Пункт")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "Жел. Дор.")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "В/Ч")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "Предназнач.")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "Наряд")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "800"
      Set colX = .Columns.add(, , "Отправлено")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "800"
      Set colX = .Columns.add(, , "Осталось")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "800"
      Set colX = .Columns.add(, , "Дата отп.")
      colX.Alignment = eLVColumnAlignCenter
      Set colX = .Columns.add(, , "Долг")
      colX.Alignment = eLVColumnAlignCenter
      colX.Width = "600"

      .ImageList(eLVSmallIcon) = img1
      .ImageList(eLVLargeIcon) = img1
      .ImageList(eLVTileImages) = img1
      .ImageList(eLVHeaderImages) = img1
      .ImageList(eLVStateImages) = img1
      .SubItemImages = True
      .BackColor = RGB(182, 252, 190)
End With
End Sub
Public Sub rodvtr_Click()
On Error Resume Next

If rodvtr.SelectedItem.Index = "1" Then
Call naryad_full("")
Else
If rodvtr.SelectedItem.Expanded = False And rodvtr.SelectedItem.Index > 1 Then
Me.Caption = "Наряд на " & rodvtr.SelectedItem.Parent & ", род войск " & rodvtr.SelectedItem
Call naryad_full("and rodv = '" & rodvtr.SelectedItem & "' and okr = '" & rodvtr.SelectedItem.Parent & "'")
End If
If rodvtr.SelectedItem.Index > 1 And rodvtr.SelectedItem.Parent = "nothing" Then
Me.Caption = "Наряд на " & rodvtr.SelectedItem.Parent & ""
Call naryad_full("and okr = '" & rodvtr.SelectedItem & "'")
End If
End If
End Sub

Private Sub rodvtr_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then Unload Me
End Sub


Private Sub naryad_full(append As String)
On Error Resume Next

listres.View = eViewDetails
listres.Visible = False
listres.CustomDraw = True
listres.AutoArrange = True
listres.OneClickActivate = True
listres.ListItems.Clear

Dim x, Y, z As Long
Dim datt() As String
Dim datt2() As String
Dim datt3() As String
    With listres.ListItems
Call mysql.query("select okr from naryad_" & nowBase & " where narid > '0' " & append & " group by okr ORDER by okr")

stt = st
datt() = dat()
For Y = 1 To stt
        Set itmX = listres.ListItems.add()
        
        itmX.BackColor = RGB(7, 172, 213)
               
        itmX.SubItems(5).Caption = datt(1, Y)
        
            Dim sFnt As New StdFont
               sFnt.Name = "Tahoma"
               sFnt.size = 12
               sFnt.Bold = True
             itmX.Font = sFnt
        
            Call mysql.query("SELECT rodv FROM naryad_" & nowBase & " WHERE okr='" & datt(1, Y) & "' " & append & " group by rodv ORDER by rodv ASC")
            sttt = st
            datt3() = dat()
            For z = 1 To sttt
                Set itmX = listres.ListItems.add()
                itmX.SubItems(5).Caption = datt3(1, z)
                itmX.BackColor = RGB(233, 230, 84)
             
             Dim sFnt2 As New StdFont
             sFnt2.Italic = True
             sFnt2.Name = "Tahoma"
             sFnt2.size = 10
             sFnt2.Bold = True
             itmX.Font = sFnt2
               
                

            Call mysql.query("SELECT `narid` as narid1,`okrkom`,`okrkom_e`,`oblkom`,`okr`,`punkt`,`doroga`,`vch`,`vrp`,`kolvo`,`datanar`,`dolg`,`prim`,`type`, `rodv`, (select sum(kolvo) from naryad_srezki_" & nowBase & " where narid = narid1)  FROM naryad_" & nowBase & " WHERE okr='" & datt(1, Y) & "' AND rodv='" & datt3(1, z) & "' " & append & " ORder by okrkom,okrkom_e,oblkom ASC")
          datt2() = dat()
          For x = 1 To st
            Set itmX = listres.ListItems.add()
                    

                                itmX.SubItems(1).Caption = datt2(1, x)
                                itmX.SubItems(2).Caption = datt2(2, x) & datt2(3, x)
                                itmX.SubItems(3).Caption = datt2(4, x)
                                itmX.SubItems(4).Caption = datt2(15, x)
                                itmX.SubItems(5).Caption = datt2(5, x)
                                itmX.SubItems(6).Caption = datt2(6, x)
                                itmX.SubItems(7).Caption = datt2(7, x)
                                itmX.SubItems(8).Caption = datt2(8, x)
                                itmX.SubItems(9).Caption = datt2(9, x)
                                itmX.SubItems(10).Caption = datt2(10, x)
                                If datt2(16, x) = "" Then datt2(16, x) = 0
                                itmX.SubItems(11).Caption = datt2(16, x)
                                itmX.SubItems(12).Caption = Int(itmX.SubItems(10).Caption) - Int(itmX.SubItems(11).Caption)
                                itmX.SubItems(13).Caption = CnvDataSqLToWin(datt2(11, x))
                                itmX.SubItems(14).Caption = datt2(12, x)
                                
                                
                                If datt2(12, x) = "0" And itmX.SubItems(12).Caption = "0" Then itmX.BackColor = &HE0E0E0
                    If datt2(14, x) = "1" Then itmX.ForeColor = &HFF&
                   If Len(datt2(13, x)) > 2 Then itmX.SubItems(2).IconIndex = 0
 
            Next x
            
            Next z
Next Y
End With


listres.Refresh
listres.Visible = True

End Sub
Private Sub rodvtr_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
    mnu_rodv_add.Visible = True
    mnu_rodv_edit.Visible = True
    mnu_rodv_del.Visible = True
    Call PopupMenu(mnu_rodv)
    mnu_rodv_add.Visible = False
    mnu_rodv_edit.Visible = False
    mnu_rodv_del.Visible = False
    End If
    
End Sub

Private Sub txtsearch_oblkom_Change()
If Len(txtsearch_oblkom) > 0 Then Call naryad_full("and oblkom like '%" & txtsearch_oblkom & "%'"): txtsearch_pred = vbNullString: txtsearch_punkt = vbNullString: txtsearch_vch = vbNullString: txtsearch_okrkom = vbNullString: txtsearch_data = vbNullString
End Sub


Private Sub txtsearch_pred_Change()
If Len(txtsearch_pred) > 0 Then Call naryad_full("and vrp like '%" & txtsearch_pred & "%'"): txtsearch_oblkom = vbNullString: txtsearch_punkt = vbNullString: txtsearch_vch = vbNullString: txtsearch_okrkom = vbNullString: txtsearch_data = vbNullString
End Sub

Private Sub txtsearch_punkt_Change()
If Len(txtsearch_punkt) > 0 Then Call naryad_full("and punkt like '%" & txtsearch_punkt & "%'"): txtsearch_oblkom = vbNullString: txtsearch_pred = vbNullString: txtsearch_vch = vbNullString: txtsearch_okrkom = vbNullString: txtsearch_data = vbNullString
End Sub


Private Sub txtsearch_vch_Change()
If Len(txtsearch_vch) > 0 Then Call naryad_full("and vch like '%" & txtsearch_vch & "%'"): txtsearch_oblkom = vbNullString: txtsearch_pred = vbNullString: txtsearch_punkt = vbNullString: txtsearch_okrkom = vbNullString: txtsearch_data = vbNullString
End Sub

Private Sub txtsearch_data_Change()
If Len(txtsearch_data) > 0 Then Call naryad_full("and datanar like '%" & txtsearch_data & "%'"): txtsearch_oblkom = vbNullString: txtsearch_pred = vbNullString: txtsearch_punkt = vbNullString: txtsearch_vch = vbNullString: txtsearch_okrkom = vbNullString
End Sub

Private Sub txtsearch_okrkom_Change()
If Len(txtsearch_okrkom) > 0 Then Call naryad_full("and okrkom like '%" & txtsearch_okrkom & "%'"): txtsearch_oblkom = vbNullString: txtsearch_pred = vbNullString: txtsearch_punkt = vbNullString: txtsearch_vch = vbNullString: txtsearch_data = vbNullString
End Sub
