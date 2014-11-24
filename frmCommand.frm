VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommand 
   BackColor       =   &H8000000A&
   Caption         =   "Команда"
   ClientHeight    =   10440
   ClientLeft      =   15
   ClientTop       =   150
   ClientWidth     =   13485
   Icon            =   "frmCommand.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10440
   ScaleMode       =   0  'User
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageListС 
      Left            =   360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCommand.frx":08CA
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCommand.frx":11A4
            Key             =   "GO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCommand.frx":1A7E
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCommand.frx":2358
            Key             =   "LOCK"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listprnik 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8599
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483641
      BackColor       =   12648384
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "УПК"
         Object.Width           =   2540
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
         Text            =   "Военкомат"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ВУС"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Директива"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Вод."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   10155
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7574
            MinWidth        =   2018
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12.12.2012"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:33"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listcommands 
      Height          =   4665
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   8229
      SortKey         =   10
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      SmallIcons      =   "ImageListС"
      ForeColor       =   0
      BackColor       =   12648384
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
         Text            =   "Обл.ком."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Предн."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Род войск"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Пункт дислок."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "В/ч"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Для пункта"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Для части"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Кол-во"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Окр.ком."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "OT_ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Lock"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComctlLib.ListView listln 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1720
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "asd"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnu_kom 
      Caption         =   "Команды"
      Begin VB.Menu mnu_com_insert 
         Caption         =   "Вставить                         [Insert]"
      End
      Begin VB.Menu mnu_com_view 
         Caption         =   "Просмотр                         [Enter]"
      End
      Begin VB.Menu mnu_com_line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_com_refresh 
         Caption         =   "Обновить"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_com_line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_com_lock 
         Caption         =   "Заблокировать команду"
         Shortcut        =   +^{F6}
      End
      Begin VB.Menu mnu_com_unlock 
         Caption         =   "Разблокировать команду"
         Shortcut        =   +^{F7}
      End
      Begin VB.Menu mnu_com_line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_next_day 
         Caption         =   "На следующий день"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_prev_day 
         Caption         =   "На предыдущий день"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_com_line09 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_com_to_excel 
         Caption         =   "в Excel"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnu_com_line04 
         Caption         =   "-"
      End
      Begin VB.Menu mnupr 
         Caption         =   "Меню Печати"
         Begin VB.Menu mnu_com_print_all 
            Caption         =   "Печать подряд"
            Shortcut        =   +^{F12}
         End
         Begin VB.Menu mnu_com_line05 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_com_print_f27 
            Caption         =   "Форма №27"
         End
         Begin VB.Menu mnu_com_print_pr_list 
            Caption         =   "Проверочный список"
         End
         Begin VB.Menu mnu_com_print_f36 
            Caption         =   "Форма №36 (ЦНП)"
         End
         Begin VB.Menu mnu_com_print_raz_ved 
            Caption         =   "Раздаточная ведомость"
         End
         Begin VB.Menu mnu_com_print_sux_pau 
            Caption         =   "Ведомость сухих пайков"
         End
         Begin VB.Menu mnu_com_print_ved_lich_nom 
            Caption         =   "Ведомость по личным номерам"
         End
         Begin VB.Menu mnu_com_print_akt_vewj 
            Caption         =   "Акт по вещевому имуществу"
         End
         Begin VB.Menu mnu_com_print_ready_400 
            Caption         =   "Готовность по приказу №400"
         End
         Begin VB.Menu mnu_com_print_ved_krugki 
            Caption         =   "Ведомость по кружкам-ложкам"
         End
         Begin VB.Menu mnu_com_print_rap_starw 
            Caption         =   "Рапорт старшего команды"
         End
         Begin VB.Menu mnu_com_print_marw_list 
            Caption         =   "Маршрутный лист"
         End
         Begin VB.Menu mnu_com_print_naksprav 
            Caption         =   "Накладные+Справки+Извещения"
         End
      End
      Begin VB.Menu mnuinfo 
         Caption         =   "Информация"
         Begin VB.Menu mnuinfocom 
            Caption         =   "О команде"
         End
      End
   End
   Begin VB.Menu mnuprniks 
      Caption         =   "Призывники"
      Begin VB.Menu mnuAdd 
         Caption         =   "Добавить       Insert"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Удалить         Delete"
      End
      Begin VB.Menu mnuToExcel 
         Caption         =   "В Excel"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuprnikinfo 
         Caption         =   "Подробно"
      End
   End
   Begin VB.Menu Print 
      Caption         =   "Печатать"
      Begin VB.Menu marsh_print 
         Caption         =   "Маршрутники"
      End
   End
End
Attribute VB_Name = "frmCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim MREC As ADODB.Recordset
Dim tmp() As String
Private myRec As API_MYSQL
Private myRec_res As API_MYSQL_RES
Private myRec_field As API_MYSQL_FIELD
Private myRec_rows As API_MYSQL_ROWS
Public d_DATE As String

Private Sub FORm_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
   
  If KeyCode = 27 Then Unload Me

End Sub

Public Sub commANDs_load()
On Error Resume Next
Dim x As Long
Dim tt As String
Dim nC As Long
Dim CntInComm As Long
If Len(CRITICAL_OPER) > 0 Then MsgBox "Подожите, пока завершится " & CRITICAL_OPER, vbExclamation, strMAIN_TITLE: Exit Sub

If listcommands.ListItems.Count > 0 Then tt = listcommands.ListItems.Item(listcommands.SelectedItem.Index).SubItems(9)
             
        DoEvents
        If Not mysql.query("SELECT naryad_" & nowBase & ".oblkom, naryad_" & nowBase & ".vrp, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, otpravka_" & nowBase & ".fORpunkt, otpravka_" & nowBase & ".fORchast, otpravka_" & nowBase & ".kolvo, naryad_" & nowBase & ".okrkom, otpravka_" & nowBase & ".otpravkaid, otpravka_" & nowBase & ".kLock FROM naryad_" & nowBase & ", otpravka_" & nowBase & " WHERE otpravka_" & nowBase & ".data like '" & dateotp & "' AND naryad_" & nowBase & ".narid = otpravka_" & nowBase & ".narid ORDER by otpravka_" & nowBase & ".otpravkaid ASC") Then
            MsgBox "Прозошла ошибка при отображении списка команд на текущую дату." & NL2 & "Возможно, произошел разрыв соединения с SQL-сервером. Для устранения этой ошибки, рекомендуется перезапустить приложение и повторить попытку подключения", vbCritical, strMAIN_TITLE
            CRITICAL_OPER = vbNullString
            Exit Sub
        End If
        
        
        
        
        listcommands.Sorted = False
        listcommands.ListItems.Clear
            For Y = 1 To st
                      If dat(11, Y) = "2" Then Set LF = listcommands.ListItems.add(, , dat(1, Y), , "GO") Else Set LF = listcommands.ListItems.add(, , dat(1, Y), , "NO")
                         For x = 1 To 10
                            LF.SubItems(x) = dat(x + 1, Y)
                            If x = 7 Then CntInComm = CntInComm + dat(x + 1, Y)
                        Next x

                        
                        If dat(11, Y) = "4" Then
                        LF.ForeColor = &H808080
                         
                            For nC = 1 To listcommands.ListItems.Item(Y).ListSubItems.Count
                                LF.ListSubItems.Item(nC).ForeColor = &H808080
                            Next nC
                        End If
             Next Y
                       
            listcommands.Sorted = True
                     

             For x = 1 To listcommands.ListItems.Count
                    If listcommands.ListItems(x).SubItems(9) = tt Then listcommands.ListItems(x).EnsureVisible: listcommands.SelectedItem.Selected = False: listcommands.ListItems(x).Selected = True: Exit For
            Next x
            listcommands.Refresh
         
    Me.Caption = " Список комманд на " & d_DATE
    
    If st > "0" Then
     p_com_id = Int(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
     Call prnik_load(0)
    End If
    Call ReSizeColumnHeaders(listcommands)


End Sub

Private Sub Form_Resize()
On Error Resume Next
listcommands.Move 0, 0, ScaleWidth, Me.Height / 2 - 1000
listln.Move 0, Me.Height / 2 - 1000, ScaleWidth, Me.Height / 8 - 300
listprnik.Move 0, Me.Height / 2, ScaleWidth, Me.Height / 2 - 1000
Call ReSizeColumnHeaders(listcommands)
Call ReSizeColumnHeaders(listprnik)
End Sub


Private Sub ListcommANDs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    listcommands.Sorted = True
    
    If listcommands.SortKey = ColumnHeader.Index - 1 Then
        If listcommands.SortOrder = lvwDescending Then listcommands.SortOrder = lvwAscending Else listcommands.SortOrder = lvwDescending
    Else
        listcommands.SortOrder = lvwAscending
        listcommands.SortKey = ColumnHeader.Index - 1
    End If
    
End Sub

Private Sub listcommANDs_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    p_com_id = Int(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
    If p_com_id > 0 Then Call prnik_load(0)
End Sub

Private Sub listcommANDs_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim x As Long
      If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
    If KeyCode = 45 Then
        If acl = "G" Or acl = "O" Then
           frmInsert.Show vbModal, Me
        Else
            MsgBox "Извините Вашему пользователю запрещен доступ на добавление команд!", vbInformation, "Доступ запрещен!"
        End If
    End If

    If KeyCode = 46 Then
        If acl = "G" Or acl = "O" Then
            Call Del_CommAND(listcommands.ListItems(listcommands.SelectedItem.Index).Text, Int(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9)))
        Else
            MsgBox "Извините Вашему пользователю запрещен доступ на удаление команд!", vbInformation, "Доступ запрещен!"
        End If
    End If
    If KeyCode = 13 Then
        p_com_id = Int(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
        If p_com_id > 0 Then Call prnik_load(0)
    End If
    If KeyCode = 120 Then Call PopupMenu(mnupr)
    
    If KeyCode = 27 Then Unload Me
    If KeyCode = 116 Then Call commANDs_load
    
    If KeyCode = 114 Then frmSearch.Show vbModal, Me
    If KeyCode = 115 Then Call commANDs_load
    If KeyCode = 38 Then
        p_com_id = Int(listcommands.ListItems(listcommands.SelectedItem.Index - 1).SubItems(9))
            If p_com_id > 0 Then Call prnik_load(0)
    End If
    If KeyCode = 40 Then
        p_com_id = Int(listcommands.ListItems(listcommands.SelectedItem.Index + 1).SubItems(9))
        If p_com_id > 0 Then Call prnik_load(0)
    End If

End Sub

Private Sub listcommANDs_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    If Button = 2 Then Call PopupMenu(mnupr)
    If Button = 4 Then Call PopupMenu(mnuinfo)
    
End Sub
Sub Del_CommAND(strCom As String, lngOtpr_Id As Long)
On Error Resume Next
   
            Dim sel As Long
                Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid = '" & lngOtpr_Id & "'")
                
                If dat(1, 1) = 0 Then
                        mysql.query ("DELETE FROM otpravka_" & nowBase & " WHERE otpravkaid =" & lngOtpr_Id)
                Else
                    MsgBox "Вы не можете удалить команду, содержащую призывников!", vbExclamation, strMAIN_TITLE
                End If
                Call log_sql("1", "8", lngOtpr_Id, strCom)
                Call commANDs_load
   
End Sub

Private Sub listprnik_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    
    listprnik.Sorted = True
    
    If listprnik.SortKey = ColumnHeader.Index - 1 Then
        If listprnik.SortOrder = lvwDescending Then listprnik.SortOrder = lvwAscending Else listprnik.SortOrder = lvwDescending
    Else
        listprnik.SortOrder = lvwAscending
        listprnik.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub listprnik_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim ss As Long
ss = Len(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
If ss > "0" Then
    If Button = 2 Then
            Call PopupMenu(mnuprniks)
    End If
End If
End Sub
Private Sub mnu_com_insert_Click()
On Error Resume Next
If acl = "G" Or acl = "O" Then frmInsert.Show vbModal, Me
End Sub
Private Sub mnu_com_lock_Click()
If acl = "G" Or acl = "O" Then
On Error Resume Next
Dim tt As String
If listcommands.ListItems.Count > 0 Then tt = listcommands.ListItems.Item(listcommands.SelectedItem.Index).SubItems(9)
            Dim x As Long
            Dim nC As Long
            Dim d_COM_OTPR As Long
            Dim kLock As Long
            
            If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
            CRITICAL_OPER = "блокировка команд"

            For x = listcommands.ListItems.Count To 1 Step -1
              If listcommands.ListItems(x).Selected Then
                    id_COM_OTPR = Int(listcommands.ListItems(x).SubItems(9))
                    kLock = 2
                    Call mysql.query("UPDATE `prnik_" & nowBase & "` set `lock`=2 WHERE `otprvid` = " & id_COM_OTPR)
                    Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kLock`=2  WHERE `otpravkaid` = " & id_COM_OTPR)
                    Call mysql.query("SELECT fam,name,otch,txtvk FROM prnik_" & nowBase & " WHERE otprvid = " & id_COM_OTPR)
                    Dim inf As String
                    inf = dat(1, 1) & " " & dat(2, 1) & " " & dat(3, 1) & " " & dat(4, 1)
                    Call log_sql("1", "5", id_COM_OTPR, "")
                  
                      listcommands.ListItems(x).SubItems(10) = "4"
                      listcommands.ListItems.Item(x).ForeColor = &H808080
                      listcommands.ListItems.Item(x).SmallIcon = "GO"
                      For nC = 1 To listcommands.ListItems.Item(x).ListSubItems.Count
                           listcommands.ListItems.Item(x).ListSubItems.Item(nC).ForeColor = &H808080
                      Next nC
                      listcommands.ListItems.Item(x).Selected = False
                      listcommands.Refresh

              End If
            Next x
            
   
    
    For x = 1 To listcommands.ListItems.Count
       If listcommands.ListItems(x).SubItems(9) = tt Then listcommands.ListItems(x).EnsureVisible: listcommands.ListItems(x).Selected = True:  Exit For
    Next x
            
    Screen.MousePointer = vbDefault
    CRITICAL_OPER = vbNullString
    End If
End Sub

Private Sub mnu_com_print_akt_krugki_Click()
Call Cnv.To_AKT_Krugki(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text, False)
End Sub

Private Sub mnu_com_print_akt_vewj_Click()
Call Cnv.To_AKT(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), False)
End Sub
' Печать подряд
Private Sub mnu_com_print_all_Click()
On Error Resume Next
'Dim auto_print As Long
'auto_print = get_config("auto_print")
progress.ProgressBar1.Max = 10
progress.Caption = "Печать команды"
frmCommand.Hide

progress.Show


Call Cnv.To_vedomostj_lichnomer(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text, True) '''''Нормально
progress.ProgressBar1 = 1
Call Cnv.To_AKT(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), True) '''''Нормально
progress.ProgressBar1 = 2
Call Cnv.To_400(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), True)
progress.ProgressBar1 = 3
Call Cnv.To_rapORt(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), True) '''''Нормально
progress.ProgressBar1 = 4
Call Cnv.To_sPayok(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), False)
progress.ProgressBar1 = 5
Call Cnv.To_vedomostj_krugki(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text, True) '''''Нормально
progress.ProgressBar1 = 6
Call Cnv.To_F8(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text) '''''
progress.ProgressBar1 = 7
Call Cnv.naksprav(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9)) '''''Нормально
progress.ProgressBar1 = 8
Call Cnv.To_F27(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9)) '''''Нормально
progress.ProgressBar1 = 9
Call Cnv.To_marw(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9)) '''''Нормально

Unload progress
frmCommand.Show
MsgBox ("Команда распечатана.")

End Sub

Private Sub mnu_com_print_f27_Click()
Call Cnv.To_F27(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
End Sub
Private Sub mnu_com_print_naksprav_Click()
Call Cnv.naksprav(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnu_com_print_marw_list_Click()
Call Cnv.To_marw(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnu_com_print_pr_list_Click()
On Error Resume Next

    Call Cnv.To_List(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnu_com_print_rap_starw_Click()
Call Cnv.To_rapORt(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), False)
End Sub

Private Sub mnu_com_print_raz_ved_Click()
Call Cnv.To_F8(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text)
End Sub

Private Sub mnu_com_print_ready_400_Click()
Call Cnv.To_400(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), False)
End Sub

Private Sub mnu_com_print_sux_pau_Click()
Call Cnv.To_sPayok(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), False)
End Sub

Private Sub mnu_com_print_ved_krugki_Click()
Call Cnv.To_vedomostj_krugki(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text, False)
End Sub

Private Sub mnu_com_print_ved_lich_nom_Click()
Call Cnv.To_vedomostj_lichnomer(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9), listcommands.ListItems(listcommands.SelectedItem.Index).Text, False)
End Sub

Private Sub mnu_com_refresh_Click()
Call commANDs_load
End Sub

Private Sub mnu_com_to_excel_Click()
Call Cnv.ResoultSearch(listcommands, True, "Список команд", Caption)
End Sub

Private Sub mnu_com_unlock_Click()
If acl = "G" Or acl = "O" Then
On Error Resume Next
Dim x As Long
Dim nC As Long



Dim tt As String
If listcommands.ListItems.Count > 0 Then tt = listcommands.ListItems.Item(listcommands.SelectedItem.Index).SubItems(9)

      If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
      CRITICAL_OPER = "разблокировка команд"
    For x = listcommands.ListItems.Count To 1 Step -1
        If listcommands.ListItems(x).Selected Then
                id_COM_OTPR = Int(listcommands.ListItems(x).SubItems(9))
                    kLock = 4
                    Call mysql.query("UPDATE `prnik_" & nowBase & "` set `lock`=0 WHERE `otprvid` = " & id_COM_OTPR)
                    Call mysql.query("UPDATE `otpravka_" & nowBase & "` SET `kLock`=0  WHERE `otpravkaid` = " & id_COM_OTPR)
                      Call mysql.query("SELECT fam,name,otch,txtvk FROM prnik_" & nowBase & " WHERE otprvid = " & id_COM_OTPR)
                    Dim inf As String
                    inf = dat(1, 1) & " " & dat(2, 1) & " " & dat(3, 1) & " " & dat(4, 1)
                    Call log_sql("1", "6", id_COM_OTPR, "")
                    
                listcommands.ListItems(x).SubItems(10) = "0"
            
                listcommands.ListItems.Item(x).SmallIcon = "NO"
                For nC = 1 To listcommands.ListItems.Item(x).ListSubItems.Count
                     listcommands.ListItems.Item(x).ListSubItems.Item(nC).ForeColor = listcommands.ForeColor
                Next nC
                listcommands.ListItems.Item(x).Selected = False
                listcommands.Refresh
               End If
    Next x

    For x = 1 To listcommands.ListItems.Count
       If listcommands.ListItems(x).SubItems(9) = tt Then listcommands.ListItems(x).EnsureVisible: listcommands.ListItems(x).Selected = True:  Exit For
    Next x
    
Screen.MousePointer = vbDefault
CRITICAL_OPER = vbNullString
End If
End Sub
Private Sub Form_Load()
'''''''''''''''''''''''''''''''''''Автопечать
Dim k As Long
    Set LF = listln.ListItems.add(, , "Личные номера")
    Set LF = listln.ListItems.add(, , "Акт")
    Set LF = listln.ListItems.add(, , "400")
    Set LF = listln.ListItems.add(, , "Рапорт")
    Set LF = listln.ListItems.add(, , "Сухой паёк")
    Set LF = listln.ListItems.add(, , "Кружки-ложки")
    Set LF = listln.ListItems.add(, , "Форма №8")
    Set LF = listln.ListItems.add(, , "Накладные, справки, извещения")
    Set LF = listln.ListItems.add(, , "Форма №27")
    If get_config("auto_print") = "1" Then
    For k = 1 To listln.ListItems.Count
    listln.ListItems(k).Checked = True
    Next k
    End If
'''''''''''''''''''''''''''''''''''''''''
                
End Sub


Private Sub mnu_next_day_Click()
On Error Resume Next
Dim datetmp() As String
frmMain.cal.NextDay
dateotp = frmMain.cal.Year & "-" & Format(frmMain.cal.Month, "00") & "-" & Format(frmMain.cal.Day, "00")
datetmp() = Split(CnvDataSqLToWin(dateotp), ".")

p_com_id = 0
d_DATE = frmMain.cal.Day & " " & MonthName(frmMain.cal.Month, False) & " " & frmMain.cal.Year
listcommands.ListItems.Clear
listprnik.ListItems.Clear

Call commANDs_load
End Sub

Private Sub mnu_prev_day_Click()
On Error Resume Next
Dim datetmp() As String
frmMain.cal.PreviousDay
dateotp = frmMain.cal.Year & "-" & Format(frmMain.cal.Month, "00") & "-" & Format(frmMain.cal.Day, "00")
datetmp() = Split(CnvDataSqLToWin(dateotp), ".")

p_com_id = 0
d_DATE = frmMain.cal.Day & " " & MonthName(frmMain.cal.Month, False) & " " & frmMain.cal.Year
listcommands.ListItems.Clear
listprnik.ListItems.Clear

Call commANDs_load
End Sub

Private Sub mnuAdd_Click()
If acl = "G" Or acl = "O" Then frmUpk.Show vbModal, Me
End Sub

Private Sub mnuLock_Click()

End Sub

Private Sub mnuprnikinfo_Click()
On Error Resume Next
Call Listprnik_DblClick
End Sub

Private Sub mnuToExcel_Click()
Call Cnv.ResoultSearch(listprnik, True, Trim$(Caption))    '
End Sub
Private Sub ListcommANDs_ColumnClicks(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    listcommands.Sorted = True
  
    If listcommands.SortKey = ColumnHeader.Index - 1 Then
        If listcommands.SortOrder = lvwDescending Then listcommands.SortOrder = lvwAscending Else listcommands.SortOrder = lvwDescending
    Else
        listcommands.SortOrder = lvwAscending
        listcommands.SortKey = ColumnHeader.Index - 1
    End If
End Sub
Private Sub Listprnik_DblClick()
On Error Resume Next

    expupk = listprnik.ListItems(listprnik.SelectedItem.Index).Text
    frmInfoPr.Show vbModal, Me
End Sub
Private Sub listprnik_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim ss As Long
 If KeyCode = 27 Then
  Unload Me
  Exit Sub
 End If
ss = Len(listcommands.ListItems(listcommands.SelectedItem.Index).SubItems(9))
If ss > "0" Then
        If KeyCode = 45 Then If acl = "G" Or acl = "O" Then frmUpk.Show vbModal, Me
        If KeyCode = 46 Then If acl = "G" Or acl = "O" Then Call del_pr
        If KeyCode = 13 Then Listprnik_DblClick
        If KeyCode = 116 Then prnik_load (0)
        If KeyCode = 17 Then listprnik.MultiSelect = True
        If KeyCode = 114 Then frmSearch.Show vbModal, Me
End If
End Sub
Function del_pr()
On Error Resume Next
If acl = "G" Or acl = "O" Then
    Dim x As Long
    Call mysql.query("SELECT `lock` FROM prnik_" & nowBase & " WHERE idprnik = " & listprnik.ListItems.Item(listprnik.SelectedItem.Index).Text)
        If CLng(dat(1, 1)) = 4 Then MsgBox "Простите, призывник находится в убывшей команде. Трогать его нельзя!", vbExclamation, strMAIN_TITLE: Exit Function
            If MsgBox("Подтверждаете удаление призывника из команды?", vbYesNo + vbQuestion, strMAIN_TITLE) = vbYes Then
                For x = listprnik.ListItems.Count To 1 Step -1
                    If listprnik.ListItems(x).Selected Then del_sql (listprnik.ListItems(x).Text): listprnik.ListItems.Remove (x)
                Next x
     End If
   End If
   commANDs_load
            listcommands.Refresh
End Function
Public Sub del_sql(strVAl As String)
On Error Resume Next
Dim komm As String
Dim otid As Long
            Call mysql.query("SELECT otprvid from prnik_" & nowBase & " where idprnik = '" & strVAl & "'")
            otid = Int(dat(1, 1))
            Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `otprvid` = '0', vus = '' WHERE `idprnik` = '" & strVAl & "'")
            Call mysql.query("SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid='" & otid & "'")
            Call mysql.query("UPDATE otpravka_" & nowBase & " set otpravka_" & nowBase & ".kolvo='" & dat(1, 1) & "' WHERE otpravka_" & nowBase & ".otpravkaid='" & otid & "'")
            Call mysql.query("SELECT otpravka_" & nowBase & ".narid from otpravka_" & nowBase & ", prnik_" & nowBase & " where otpravka_" & nowBase & ".otpravkaid = prnik_" & nowBase & ".otprvid and prnik_" & nowBase & ".idprnik = '" & strVAl & "'")
            komm = Int(dat(1, 1))
 Call log_sql("0", "1", strVAl, komm)
End Sub

Public Sub prnik_load(oper As Integer)
On Error Resume Next
        Dim tCol As Long
        Dim tSel As Long
        Dim tt As String
        If Len(CRITICAL_OPER) > 0 Then MsgBox "Подожите, пока завершится " & CRITICAL_OPER, vbExclamation, strMAIN_TITLE: Exit Sub
        If listprnik.ListItems.Count > 0 Then tt = listprnik.ListItems.Item(listprnik.SelectedItem.Index).Text
       
        CRITICAL_OPER = "обновление списка призывников"
      '  'Screen.MousePointer = vbHourglass
        SB1.Panels(1).Text = "Чтение данных..."
        SB1.Refresh

        Call mysql.query("SELECT `idprnik`, `fam`, `name`, `otch`, `txtvk`, `vus`, `dir`, `vod`, `lock` FROM prnik_" & nowBase & " WHERE otprvid = '" & p_com_id & "' ORder by txtvk,fam ASC")
        CRITICAL_OPER = "отображение списка призывников"
        SB1.Panels(1).Text = "Отображение..."
        SB1.Refresh
        
                
        
           Dim c As Long
           INP_UPK = VAl(INP_UPK)
                    If oper = 1 Then
                             For c = 1 To listprnik.ListItems.Count
                                    If VAl(listprnik.ListItems(c).Text) = INP_UPK Then
                                            SB1.Panels(1).Text = "Всего: " & listprnik.ListItems.Count
                                            Screen.MousePointer = vbDefault
                                            MsgBox "Призывник в этой команде с номером УПК " & INP_UPK & " уже существует!", vbCritical, "": Exit Sub
                                    End If
                              Next c
                     End If
                     
            listprnik.ListItems.Clear
            listprnik.Sorted = False
            
                For s = 1 To st
                        If dat(9, s) = "0" Then Set LF = listprnik.ListItems.add(, , dat(1, s), , "OK") Else Set LF = listprnik.ListItems.add(, , dat(1, s), , "LOCK")
                            For c = 2 To 7
                                    LF.SubItems(c - 1) = dat(c, s)
                            Next c
                            LF.SubItems(7) = choose(dat(8, s) + 1, vbNullString, strCHAR_VOD)
                Next s
                
       
        listprnik.Refresh
        
        Call mysql.query("SELECT sum(vod) FROM prnik_" & nowBase & " WHERE otprvid = " & p_com_id)

        
        CRITICAL_OPER = vbNullString
        INP_UPK = 0
        SB1.Panels(1).Text = "Всего: " & listprnik.ListItems.Count
        SB1.Panels(2).Text = "Водителей: " & dat(1, 1)
        Call mysql.query("SELECT count(*) FROM prnik_" & nowBase & ",otpravka_" & nowBase & " WHERE data like '" & dateotp & "%' AND prnik_" & nowBase & ".otprvid=otpravka_" & nowBase & ".otpravkaid")
        SB1.Panels(3).Text = "В командах всего: " & dat(1, 1)
        Call mysql.query("SELECT count(*) FROM otpravka_" & nowBase & " WHERE data='" & dateotp & "'")
        SB1.Panels(4).Text = "Команд за день: " & dat(1, 1)
        Screen.MousePointer = vbDefault
       Call ReSizeColumnHeaders(listprnik)
        
        
End Sub

Private Sub mnudel_Click()
On Error Resume Next
    If acl = "G" Or acl = "O" Then Call del_pr
End Sub

Private Sub marsh_print_click()
On Error Resume Next
    Dim numb As Integer
numb = Int(VAl(InputBox("Кол-во")))
oExcelApp.Application.DisplayAlerts = False
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")

        If Err = 429 Then MsgBox "Application Microsoft Excel is not installed!", vbExclamation, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Sub

    Dim sFile As String
    sFile = sCNV_txtDirShabl & "marw.xls"
    If Not ExFile(sFile) Then MsgBox "Файл шаблона" & NL2 & sFile & NL2 & "не найден. Проверьте правильность указания пути шаблонов." & vbNewLine & "Для восстановления шаблонов воспользуйтесь вкладкой `Воссановление` в меню `Опции`", vbCritical, strMAIN_TITLE: Screen.MousePointer = vbDefault: Exit Sub
    
oExcelApp.Visible = False
    
    
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(2)



If numb > 0 Then
oWs.PrintOut Copies:=numb
        oWb.Close
        oExcelApp.Quit
End If


End Sub


