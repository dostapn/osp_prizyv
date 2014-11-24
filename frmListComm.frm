VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListComm 
   Caption         =   "Список комманд"
   ClientHeight    =   12585
   ClientLeft      =   -15
   ClientTop       =   450
   ClientWidth     =   14940
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListComm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmListComm.frx":5F32
   ScaleHeight     =   12585
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageListС 
      Left            =   1665
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListComm.frx":249E3
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListComm.frx":24F7D
            Key             =   "GO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Привязать вниз
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   12300
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "В командах:"
            TextSave        =   "В командах:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8308
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "23.04.2007"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:55"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListFiles 
      Height          =   9165
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   16166
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      SmallIcons      =   "ImageListС"
      ForeColor       =   4210752
      BackColor       =   12965598
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Предн."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Род войск"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Пункт дислок."
         Object.Width           =   1764
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
      Picture         =   "frmListComm.frx":25517
   End
   Begin VB.Menu mnuListComm 
      Caption         =   "Список команд"
      Begin VB.Menu mnuAddCom 
         Caption         =   "Добавить                              [Insert]"
      End
      Begin VB.Menu mnuView 
         Caption         =   "Просмотр                               [Enter]"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Обновить                               [F5]"
      End
      Begin VB.Menu line003 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockCommand 
         Caption         =   "Заблокировать"
         Shortcut        =   +^{F6}
      End
      Begin VB.Menu mnuUnLockCommand 
         Caption         =   "Разблокировать"
         Shortcut        =   +^{F7}
      End
      Begin VB.Menu line004 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToExcel 
         Caption         =   "В Excel"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu lene03 
         Caption         =   "-"
      End
      Begin VB.Menu print_all_a4 
         Caption         =   "Печать подряд(А4)"
      End
      Begin VB.Menu mnuPr 
         Caption         =   "Печать                                   [F9]"
         Begin VB.Menu mnuF27 
            Caption         =   "Форма №27"
         End
         Begin VB.Menu mnuList 
            Caption         =   "Проверочный список"
         End
         Begin VB.Menu mnuCNP 
            Caption         =   "Целенаправленный призыв"
         End
         Begin VB.Menu mnuRazVed 
            Caption         =   "Раздаточная ведомость"
         End
         Begin VB.Menu mnusPayok 
            Caption         =   "Ведомость сух.пайков"
         End
         Begin VB.Menu mnuAKT 
            Caption         =   "Акт по вещевому имуществу"
         End
         Begin VB.Menu p400 
            Caption         =   "Готовность по 400 приказу"
         End
         Begin VB.Menu akt_t 
            Caption         =   "Акт кружек-ложек"
         End
         Begin VB.Menu dd 
            Caption         =   "Ведомость кружек-ложек"
         End
         Begin VB.Menu raport 
            Caption         =   "Рапорт старшего команды"
         End
         Begin VB.Menu marw 
            Caption         =   "Маршрутный лист"
         End
      End
   End
End
Attribute VB_Name = "frmListComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MREC As ADODB.Recordset
Dim tmp() As String
Private myRec As API_MYSQL
Private myRec_res As API_MYSQL_RES
Private myRec_field As API_MYSQL_FIELD
Private myRec_rows As API_MYSQL_ROWS
Public d_DATE As String

Private Sub cmdAdd_Click()
On Error Resume Next
        For c = 1 To st
        
            Set LF = ListFiles.ListItems.Add(, , vbNullString)
            
                For R = 1 To col
                        LF.SubItems(R + 1) = DAT(R, c)
                Next R
        
        Next c

End Sub

Private Sub Command2_Click()
On Error Resume Next
        For c = 1 To COL_ID
        
        Call mysql.query("select oblkom from naryad where narid='5'")
        Set LF = ListFiles.ListItems.Add(, , DAT(1, 1))
        
            For R = 1 To 9
            
             LF.SubItems(R) = DAT(1, 1)
             
            Next R
        Next c
End Sub


Public Sub cmdShow()
On Error Resume Next
Dim X As Long
Dim tt As String
Dim nC As Long
Dim CntInComm As Long
Screen.MousePointer = vbHourglass
        If Len(CRITICAL_OPER) > 0 Then MsgBox "Подожите, пока завершится " & CRITICAL_OPER, vbExclamation, strMAIN_TITLE: Exit Sub

If ListFiles.ListItems.Count > 0 Then tt = ListFiles.ListItems.Item(ListFiles.SelectedItem.Index).SubItems(9)

        CRITICAL_OPER = "чтение списка призывников"
        SB1.Panels(1).Text = "Чтение данных..."
        SB1.Refresh
        DoEvents
        If Not mysql.query("select naryad.oblkom, naryad.vrp, naryad.rodv, naryad.punkt, naryad.vch, otpravka.forpunkt, otpravka.forchast, otpravka.kolvo, naryad.okrkom, otpravka.otpravkaid, otpravka.kLock from naryad, otpravka where otpravka.data like '" & strDATE & "' and naryad.narid = otpravka.narid") Then
            MsgBox "Прозошла ошибка при отображении списка команд на текущую дату." & NL2 & "Возможно, произошел разрыв соединения с SQL-сервером. Для устранения этой ошибки, рекомендуется перезапустить приложение и повторить попытку подключения", vbCritical, strMAIN_TITLE
            CRITICAL_OPER = vbNullString
            Exit Sub
        End If
        CRITICAL_OPER = "обновление списка призывников"
        SB1.Panels(1).Text = "Обновление..."
        SB1.Refresh
        
        
        
        ListFiles.Sorted = False
        ListFiles.ListItems.Clear
            For Y = 1 To st
            
                If DAT(11, Y) = "4" Then Set LF = ListFiles.ListItems.Add(, , DAT(1, Y), , "GO") Else Set LF = ListFiles.ListItems.Add(, , DAT(1, Y), , "NO")
                
                        For X = 1 To 10
                            LF.SubItems(X) = DAT(X + 1, Y)
                            If X = 7 Then CntInComm = CntInComm + DAT(X + 1, Y)
                        Next X

                        
                        If DAT(11, Y) = "4" Then
                        LF.ForeColor = &H808080
                            For nC = 1 To ListFiles.ListItems.Item(Y).ListSubItems.Count
                                LF.ListSubItems.Item(nC).ForeColor = &H808080
                            Next nC
                        End If

                
            Next Y
                        
            ListFiles.Sorted = True
           
            'If Not Len(tt) = 0 Then If ListFiles.ListItems.Count > 0 Then ListFiles.SelectedItem.Selected = False

             For X = 1 To ListFiles.ListItems.Count
                    If ListFiles.ListItems(X).SubItems(9) = tt Then ListFiles.ListItems(X).EnsureVisible: ListFiles.SelectedItem.Selected = False: ListFiles.ListItems(X).Selected = True: Exit For
            Next X



            ListFiles.Refresh
         Call mysql.query("select count(idprnik) from prnik;")
             If iCOMM_chFillSelComm = 1 Then ListFiles.SelectedItem.Selected = False

         CRITICAL_OPER = vbNullString
        SB1.Panels(1).Text = "Всего в базе: " & DAT(1, 1)
        SB1.Panels(2).Text = "Команд: " & ListFiles.ListItems.Count
        SB1.Panels(3).Text = "В командах: " & CntInComm
        CRITICAL_OPER = vbNullString
        Screen.MousePointer = vbDefault
    Me.Caption = "Список комманд на " & d_DATE
End Sub


Private Sub akt_t_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
    COM_OTP_ID = ListFiles.ListItems(ListFiles.SelectedItem.Index).Text
    Call Cnv.To_AKT_Krugki(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text, False)
End Sub

Private Sub dd_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
    COM_OTP_ID = ListFiles.ListItems(ListFiles.SelectedItem.Index).Text
    Call Cnv.To_vedomostj_krugki(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text, False)
    End Sub

Private Sub Form_Load()
    On Error Resume Next
    Caption = "Чтение списка команд..."
    Screen.MousePointer = vbHourglass
    Call cmdShow
    Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListFiles_Click()
     If iCOMM_chFillSelComm = 1 Then ListFiles.SelectedItem.Selected = False
End Sub

Private Sub ListFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    ListFiles.Sorted = True
    
    If ListFiles.SortKey = ColumnHeader.Index - 1 Then
        If ListFiles.SortOrder = lvwDescending Then ListFiles.SortOrder = lvwAscending Else ListFiles.SortOrder = lvwDescending
    Else
        ListFiles.SortOrder = lvwAscending
        ListFiles.SortKey = ColumnHeader.Index - 1
    End If
    
End Sub

Private Sub ListFiles_DblClick()
On Error Resume Next
        If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
        
        If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
        
        p_com_id = Int(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
        frmCommand.Caption = ListFiles.ListItems(ListFiles.SelectedItem.Index).Text
         frmCommand.Show vbModal, Me
        
End Sub

Private Sub listfiles_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim X As Long
      If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
    If KeyCode = 45 Then
        If acl = "G" Or acl = "O" Then
           frmInsert.Show vbModal, Me
        Else
            MsgBox "Извините Вашему пользователю запрещен доступ!", vbInformation, "Доступ запрещен!"
        End If
    End If
    If KeyCode = 46 Then
        If acl = "G" Or acl = "O" Then
            Call Del_Command(ListFiles.ListItems(ListFiles.SelectedItem.Index).Text, Int(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9)))
        Else
            MsgBox "Извините Вашему пользователю запрещен доступ!", vbInformation, "Доступ запрещен!"
        End If
    End If
    If KeyCode = 120 Then Call PopupMenu(mnupr)
    If KeyCode = 13 Then ListFiles_DblClick
    If KeyCode = 17 Then ListFiles.MultiSelect = True
    If KeyCode = 27 Then Unload Me
    If KeyCode = 116 Then Call cmdShow
    If KeyCode = 114 Then frmSearch.Show vbModal, Me
    If KeyCode = 115 Then
    Call cmdShow
    End If
    

End Sub

Sub Del_Command(strCom As String, lngOtpr_Id As Long)
On Error Resume Next
   
            Dim sel As Long
                Call mysql.query("select Count(idprnik) from prnik where otprvid = '" & lngOtpr_Id & "'")
                
                If DAT(1, 1) = 0 Then
                        mysql.query ("DELETE from otpravka where otpravkaid =" & lngOtpr_Id)
                        mysql.query ("delete from knotp where isx=" & lngOtpr_Id)
                        If ListFiles.SelectedItem.Index = ListFiles.ListItems.Count Then
                            ListFiles.ListItems.Remove (ListFiles.SelectedItem.Index)
                            If ListFiles.ListItems.Count > 0 Then ListFiles.ListItems(ListFiles.ListItems.Count).Selected = True
                        Else
                            sel = ListFiles.SelectedItem.Index
                            ListFiles.ListItems.Remove (ListFiles.SelectedItem.Index)
                           If ListFiles.ListItems.Count > 0 Then ListFiles.ListItems(sel).Selected = True
                        End If
                Else
                    MsgBox "Вы не можете удалить команду, содержащую призывников!", vbExclamation, strMAIN_TITLE
                End If
                Call log_sql(lgn, "Удалил команду - " & strCom)
   
End Sub

Private Sub ListFiles_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
     If iCOMM_chFillSelComm = 1 Then ListFiles.SelectedItem.Selected = False
    If KeyCode = 17 Then ListFiles.MultiSelect = False
End Sub

Private Sub listfiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If Button = 2 Then
            Call PopupMenu(mnuListComm)
    End If
    
End Sub

Private Sub marw_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
   Call Cnv.To_marw(listcommands.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnuAddCom_Click()
    On Error Resume Next
    
  If acl = "G" Or acl = "O" Then frmInsert.Show vbModal, Me
    
End Sub

Private Sub mnuAKT_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
   Call Cnv.To_AKT(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), False)
End Sub

Private Sub mnuCNP_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
Call SaveSetting("RNB_PRIZ_DB", "Settings", "tmpOblKom", ListFiles.ListItems(ListFiles.SelectedItem.Index).Text)
    Call Cnv.To_F36(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnuList_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
    Call Cnv.To_List(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnuLockCommand_Click()
If acl = "G" Or acl = "O" Then
On Error Resume Next
Dim tt As String
If ListFiles.ListItems.Count > 0 Then tt = ListFiles.ListItems.Item(ListFiles.SelectedItem.Index).SubItems(9)
            Dim X As Long
            Dim nC As Long
            Dim d_COM_OTPR As Long
            Dim kLock As Long
            If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
            If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
            CRITICAL_OPER = "блокировка команд"
            Screen.MousePointer = vbHourglass
            For X = ListFiles.ListItems.Count To 1 Step -1
              If ListFiles.ListItems(X).Selected Then
                    id_COM_OTPR = Int(ListFiles.ListItems(X).SubItems(9))
                    kLock = 4
                    Call mysql.query("update `prnik` set `lock`=4 where `otprvid` = " & id_COM_OTPR)
                    Call mysql.query("UPDATE `otpravka` SET `kLock`=4  WHERE `otpravkaid` = " & id_COM_OTPR)
                    Call mysql.query("select fam,name,otch,txtvk from prnik where otprvid = " & id_COM_OTPR)
                    Dim inf As String
                    inf = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1)
                    Call log_sql(lgn, "Блокировал призывника " & inf & " военкомат")
                  
                      ListFiles.ListItems(X).SubItems(10) = "4"
                      ListFiles.ListItems.Item(X).ForeColor = &H808080
                      ListFiles.ListItems.Item(X).SmallIcon = "GO"
                      For nC = 1 To ListFiles.ListItems.Item(X).ListSubItems.Count
                           ListFiles.ListItems.Item(X).ListSubItems.Item(nC).ForeColor = &H808080
                      Next nC
                      ListFiles.ListItems.Item(X).Selected = False
'                      ListFiles.Refresh
                 ' End If
              End If
            Next X
            
   
    
    For X = 1 To ListFiles.ListItems.Count
       If ListFiles.ListItems(X).SubItems(9) = tt Then ListFiles.ListItems(X).EnsureVisible: ListFiles.ListItems(X).Selected = True:  Exit For
    Next X
            
    Screen.MousePointer = vbDefault
    CRITICAL_OPER = vbNullString
    End If
End Sub

Private Sub mnuRazVed_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
    COM_OTP_ID = ListFiles.ListItems(ListFiles.SelectedItem.Index).Text
    Call Cnv.To_F8(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text)
End Sub

Private Sub mnurefresh_Click()
On Error Resume Next
Call cmdShow
End Sub
Private Sub mnuF27_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
Call SaveSetting("RNB_PRIZ_DB", "Settings", "tmpOblKom", ListFiles.ListItems(ListFiles.SelectedItem.Index).Text)
    Call Cnv.To_F27(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnusPayok_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
    'Call Cnv.To_sPayok(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
End Sub

Private Sub mnuToExcel_Click()
    Call Cnv.ResoultSearch(ListFiles, True, "ListCommand", Caption)
End Sub

Private Sub mnuUnLockCommand_Click()
If acl = "G" Or acl = "O" Then
On Error Resume Next
Dim X As Long
Dim nC As Long

If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub

Dim tt As String
If ListFiles.ListItems.Count > 0 Then tt = ListFiles.ListItems.Item(ListFiles.SelectedItem.Index).SubItems(9)


      If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
      CRITICAL_OPER = "разблокировка команд"
      

Screen.MousePointer = vbHourglass
    For X = ListFiles.ListItems.Count To 1 Step -1
        If ListFiles.ListItems(X).Selected Then
            'If Lock_Comm(Int(ListFiles.ListItems(x).SubItems(9)), 0) Then
                id_COM_OTPR = Int(ListFiles.ListItems(X).SubItems(9))
                    kLock = 4
                    Call mysql.query("update `prnik` set `lock`=0 where `otprvid` = " & id_COM_OTPR)
                    Call mysql.query("UPDATE `otpravka` SET `kLock`=0  WHERE `otpravkaid` = " & id_COM_OTPR)
                      Call mysql.query("select fam,name,otch,txtvk from prnik where otprvid = " & id_COM_OTPR)
                    Dim inf As String
                    inf = DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & " " & DAT(4, 1)
                    Call log_sql(lgn, "Разблокировал призывника " & inf & " военкомат")
            ListFiles.ListItems(X).SubItems(10) = "0"
                ListFiles.ListItems.Item(X).ForeColor = ListFiles.ForeColor
                ListFiles.ListItems.Item(X).SmallIcon = "NO"
                For nC = 1 To ListFiles.ListItems.Item(X).ListSubItems.Count
                     ListFiles.ListItems.Item(X).ListSubItems.Item(nC).ForeColor = ListFiles.ForeColor
                Next nC
                ListFiles.ListItems.Item(X).Selected = False
'                ListFiles.Refresh
            'End If
        End If
    Next X

    For X = 1 To ListFiles.ListItems.Count
       If ListFiles.ListItems(X).SubItems(9) = tt Then ListFiles.ListItems(X).EnsureVisible: ListFiles.ListItems(X).Selected = True:  Exit For
    Next X
    
Screen.MousePointer = vbDefault
CRITICAL_OPER = vbNullString
End If
End Sub

Private Sub mnuView_Click()
    ListFiles_DblClick
End Sub

Private Sub p400_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
   Call Cnv.To_400(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), False)
End Sub

Private Sub print_all_a4_Click()
Call Cnv.To_AKT(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), True)
Call Cnv.To_AKT_Krugki(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text, True)
Call Cnv.To_400(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), True)
Call Cnv.To_vedomostj_krugki(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text, True)
Call Cnv.To_raport(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), True)
Call Cnv.To_sPayok(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), True)
Call Cnv.To_F27(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
'/--------
Call mysql.query("select naryad.vch, otpravka.narid from otpravka, naryad where otpravka.otpravkaid='" & ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9) & "' and naryad.narid=otpravka.narid")
If Right$(DAT(1, 1), 3) = "ЦНП" Then Call Cnv.To_F36(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))


'/--------
Call Cnv.To_marw(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9))
COM_OTP_ID = ListFiles.ListItems(ListFiles.SelectedItem.Index).Text
Call Cnv.To_F8(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), ListFiles.ListItems(ListFiles.SelectedItem.Index).Text)
End Sub

Private Sub raport_Click()
On Error Resume Next
If ListFiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
'Call SaveSetting("RNB_PRIZ_DB", "Settings", "tmpOblKom", ListFiles.ListItems(ListFiles.SelectedItem.Index).Text)
   Call Cnv.To_raport(ListFiles.ListItems(ListFiles.SelectedItem.Index).SubItems(9), False)
'Call Cnv.To_raport
End Sub
