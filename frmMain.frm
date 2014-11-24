VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   8055
   ClientTop       =   6810
   ClientWidth     =   7950
   DrawWidth       =   2
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
   ForeColor       =   &H000000FF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000E&
      Caption         =   "     Жетоны                Прод."
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   2160
      Width           =   2655
      Begin VB.Label prod 
         BackColor       =   &H8000000E&
         Caption         =   "Label2"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Ln 
         BackColor       =   &H8000000E&
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Текущая база"
      Height          =   1095
      Left            =   5280
      TabIndex        =   4
      Top             =   0
      Width           =   2655
      Begin VB.Label now_base 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Пользователь"
      Height          =   1335
      Left            =   5280
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
      Begin VB.Label now_user 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Выбор базы"
      Height          =   1095
      Left            =   5280
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   15
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Команды по дате"
      Height          =   3975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin MSACAL.Calendar cal 
         Height          =   2655
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   4455
         _Version        =   524288
         _ExtentX        =   7858
         _ExtentY        =   4683
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2007
         Month           =   5
         Day             =   3
         DayLength       =   1
         MonthLength     =   0
         DayFontColor    =   255
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   12582912
         GridLinesColor  =   16777215
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3945
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1720
            MinWidth        =   1720
            Text            =   "Привезли"
            TextSave        =   "Привезли"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   794
            MinWidth        =   794
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1852
            MinWidth        =   1852
            Text            =   "В командах"
            TextSave        =   "В командах"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   794
            MinWidth        =   794
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2514
            MinWidth        =   2514
            Text            =   "Осталось всего"
            TextSave        =   "Осталось всего"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   926
            MinWidth        =   926
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1852
            MinWidth        =   1852
            TextSave        =   "12.12.2012"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:41"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman CYR"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1800
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   840
      Top             =   7320
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1020
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   88
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_prizyvnik 
      Caption         =   "Призывник"
      Begin VB.Menu mnuSearch 
         Caption         =   "Найти"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuDir 
         Caption         =   "Директивы"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Сервис"
      Begin VB.Menu mnuReport 
         Caption         =   "Отчет за день"
         Shortcut        =   ^R
      End
      Begin VB.Menu report_to_date 
         Caption         =   "Отчет за дату"
         Shortcut        =   ^D
      End
      Begin VB.Menu line00111 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_nar 
         Caption         =   "Наряд"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_service_spiski_ot_komp 
         Caption         =   "Списки от компл. "
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu top_menu_04 
      Caption         =   "Статистика"
      Begin VB.Menu mnu_stat_rodv 
         Caption         =   "По Роду Войск"
      End
      Begin VB.Menu mnu_line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_stat_obraz 
         Caption         =   "Образование"
         Begin VB.Menu mnu_stat_obraz_fam 
            Caption         =   "По фамильно"
            Begin VB.Menu mnu_stat_obraz_vsh_fam 
               Caption         =   "Высшее"
            End
            Begin VB.Menu mnu_stat_obraz_nvsh_fam 
               Caption         =   "Незаконченное высшее"
            End
            Begin VB.Menu mnu_stat_obraz_srspec_fam 
               Caption         =   "Средне-специальное"
            End
            Begin VB.Menu mnu_stat_obraz_sr_fam 
               Caption         =   "Среднее"
            End
            Begin VB.Menu mnu_stat_obraz_nsr_fam 
               Caption         =   "Неполное среднее"
            End
            Begin VB.Menu mnu_stat_obraz_9k_fam 
               Caption         =   "9 классов и ниже"
            End
         End
         Begin VB.Menu mnu_stat_obraz_vk 
            Caption         =   "По военокомату"
            Begin VB.Menu mnu_stat_obraz_vsh_vk 
               Caption         =   "Высшее"
            End
            Begin VB.Menu mnu_stat_obraz_nvsh_vk 
               Caption         =   "Незаконченное высшее"
            End
            Begin VB.Menu mnu_stat_obraz_srspec_vk 
               Caption         =   "Средне-специальное"
            End
            Begin VB.Menu mnu_stat_obraz_sr_vk 
               Caption         =   "Среднее"
            End
            Begin VB.Menu mnu_stat_obraz_nsr_vk 
               Caption         =   "Неполное среднее"
            End
            Begin VB.Menu mnu_stat_obraz_9k_vk 
               Caption         =   "9 классов и ниже"
            End
         End
      End
      Begin VB.Menu mnu_line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_stat_date_privoz 
         Caption         =   "По Датам Привоза"
      End
      Begin VB.Menu mnu_stat_date_otp 
         Caption         =   "По Датам Отправки"
      End
      Begin VB.Menu mnu_line03 
         Caption         =   "-"
      End
      Begin VB.Menu godr 
         Caption         =   "По Годам"
      End
   End
   Begin VB.Menu top_menu_05 
      Caption         =   "Документы"
      Begin VB.Menu com_r 
         Caption         =   "Книга учета команд"
      End
      Begin VB.Menu mnuListsVK 
         Caption         =   "Генерировать списки"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu genalltoof 
         Caption         =   "Генерировать списки в один файл"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu gen01 
         Caption         =   "Генерировать списки c предназначением"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu line001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_service_vipa_d 
         Caption         =   "Випа Д"
      End
      Begin VB.Menu mnu_service_vipa_p 
         Caption         =   "Випа Р"
      End
   End
   Begin VB.Menu mysql_f 
      Caption         =   "БД"
      Begin VB.Menu optimiz 
         Caption         =   "Оптимизация таблиц"
      End
      Begin VB.Menu mnu_table_check 
         Caption         =   "Проверка таблиц"
      End
      Begin VB.Menu linem01 
         Caption         =   "-"
      End
      Begin VB.Menu create_db_new 
         Caption         =   "Создать базу"
      End
   End
   Begin VB.Menu mnu_system 
      Caption         =   "Система"
      Begin VB.Menu mnu_system_settings 
         Caption         =   "Настройки программы"
      End
      Begin VB.Menu mnu_settings_print 
         Caption         =   "Настройки печати"
      End
      Begin VB.Menu mnu_log 
         Caption         =   "Системные сообщения"
      End
      Begin VB.Menu mnu_system_users 
         Caption         =   "Пользователи"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sShow As Boolean
Option Explicit

Private Sub cmdDiscon_Click()
    Set mysql = Nothing
End Sub

Private Sub cmdRCon_Click()
On Error GoTo errhANDler
        'Screen.MousePointer = vbHourglass
        
        Caption = "Соединяемся с " & txtHost & " [" & txtDB & "]..."
        DoEvents
        Set mysql = New cMysql
        mysql.real_connect txtHost, txtUser, txtpass, txtDB, CLng(VAl(txtPORt)), , 0
        
        
        Exit Sub
        
errhANDler:

        Screen.MousePointer = vbDefault
         
        Caption = Err.Description
End Sub

'0


Private Sub bd_make_Click()
Dim db_old As String

db_old = InputBox("Как сохранить текущую базу:", "Создание новой базы", "prizyv_")
Call mysql.query("")

End Sub

Private Sub mnu_stat_date_privoz_Click()
Call Cnv.stat_date
End Sub
Private Sub Cal_DblClick()

Dim datetmp() As String
dateotp = cal.Year & "-" & Format(cal.Month, "00") & "-" & Format(cal.Day, "00")
datetmp() = Split(CnvDataSqLToWin(dateotp), ".")

p_com_id = 0
frmCommand.d_DATE = cal.Day & " " & MonthName(cal.Month, False) & " " & cal.Year
frmCommand.listcommands.ListItems.Clear
frmCommand.listprnik.ListItems.Clear
frmCommand.commANDs_load


frmCommand.Show
End Sub

Private Sub Cal_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then Cal_DblClick
End Sub

Private Sub com_r_Click()
Call Cnv.To_knotp
End Sub

Private Sub Combo1_LostFocus()
Combo1.Enabled = False
End Sub

Private Sub CommAND1_Click()
    If MsgBox("Завершить работы программы?", vbQuestion + vbOKCancel, strMAIN_TITLE) = vbOK Then Unload Me
End Sub

Private Sub create_db_new_Click()
On Error Resume Next
If acl = "G" Then
If MsgBox("Вы уверены, что хотите создать новую базу?", vbYesNo + vbQuestion, "Создание базы") = vbYes Then
    Dim last As String
    Dim pr As String
    Dim newBase As String
    Dim name_newBase As String
    Dim id_new As Long
    Call mysql.query("SELECT max(id) FROM bases")
    id_new = dat(1, 1) + 1
    Call mysql.query("SELECT VAl FROM bases WHERE id='" & dat(1, 1) & "'")
    last = dat(1, 1)
    If Right$(last, 1) = "1" Then
        newBase = Left$(last, 2) & "2"
        name_newBase = "Весна 20" & Left$(last, 2)
    End If
    If Right$(last, 1) = "2" Then
        pr = Left$(last, 2) + 1
        If pr < 10 Then pr = "0" & pr
        name_newBase = "Осень 20" & pr
        newBase = pr & "1"
    End If
'создание информации в базе BASES
Call mysql.query("insert into bases VAlues ('" & id_new & "','" & name_newBase & "','" & newBase & "','') ")
    'Создание новых таблиц
        Call mysql.query("show tables")
        datt() = dat()
        For x = 1 To st
            If Right(datt(1, x), 3) = nowBase Then
                Call mysql.query("create table " & Left(datt(1, x), Len(datt(1, x)) - 4) & "_" & newBase & " select * from " & datt(1, x))
                Call mysql.query("Truncate table `" & Left(datt(1, x), Len(datt(1, x)) - 4) & "_" & newBase & "`")
            End If
        Next x
        
        MsgBox "Новая база " & name_newBase & " создана!", vbExclamation, "Создание новой БД"
    End If
End If
End Sub
Private Sub combo1_Click()
Combo1.Enabled = True
Call mysql.query("SELECT VAl FROM bases WHERE name='" & Combo1.Text & "'")
         nowBase = dat(1, 1)
now_base.Caption = Combo1.Text
Call priziv_info
Me.Caption = "База: " & Combo1.Text
               
End Sub



Private Sub Form_Load()
Dim x As Long
Dim strr As String
    Dim VAldata As Integer
      Call INIT_VRP_LIST
          
      
      Me.Refresh
      DoEvents
     
      Show

    Combo1.Enabled = False
    Call mysql.query("SELECT name FROM bases ORder by id ASC")
    For x = 1 To st
    Combo1.AddItem (dat(1, x))
    Next x
    Call mysql.query("SELECT max(id) FROM bases")
    Dim basesel As Long
    basesel = dat(1, 1)
    Combo1.ListIndex = basesel - 1
    Caption = "База: " & Combo1.Text
    Call mysql.query("SELECT VAl FROM bases WHERE id='" & basesel & "'")
    nowBase = dat(1, 1)
      Me.Enabled = True
    Call get_access
     info_panel
    Call mysql.query("SELECT fio FROM users WHERE name = '" & lgn & "'")
        now_user.Caption = dat(1, 1)
    Call priziv_info
    Call nom_prikaz

    'Call ip_upd
    

    End Sub

Public Sub priziv_info()
    On Error Resume Next
    
        Call mysql.query("SELECT max(id) FROM bases")
        Call mysql.query("SELECT VAl FROM bases WHERE id='" & dat(1, 1) & "'")
        If nowBase = dat(1, 1) Then
        cal.VAlue = Date
        Else
            If Right$(nowBase, 1) = "1" Then
                cal.Month = 5
                cal.Day = 1
                cal.Year = "20" & Left$(nowBase, 2)
            Else
       
                cal.Month = 11
                cal.Day = 1
                cal.Year = "20" & Left$(nowBase, 2)
            End If
        
        End If
   
   
 
   
   
   
End Sub
Public Sub nom_prikaz()
        Call mysql.query("Select count(*) from `dayprikaz` where data = '" & Format(Now, "YYYY-MM-DD") & "'")
If dat(1, 1) > 0 Then
    Call mysql.query("SELECT ln FROM dayprikaz WHERE data = '" & Format(Now, "YYYY-MM-DD") & "'")
    Ln.Caption = dat(1, 1)
    Call mysql.query("SELECT prod FROM dayprikaz WHERE data = '" & Format(Now, "YYYY-MM-DD") & "'")
    prod.Caption = dat(1, 1)
    Else
    Frame6.BackColor = &H80FFFF
    Ln.BackColor = &H80FFFF
    prod.BackColor = &H80FFFF
    Ln.Caption = "Нет"
    Ln.ForeColor = &HFF&
    prod.Caption = "Нет"
    prod.ForeColor = &HFF&
    End If
   
End Sub
Private Sub info_panel()
 
 Dim ddate As String
      ddate = Date
     ' Label4.Caption = Date
Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `dataosp` = '" & CnvDataWinToSql(ddate) & "'")
    STB.Panels(2).Text = dat(1, 1)
Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `otprvid` <> 0 AND prnik_" & nowBase & ".dataosp = '" & CnvDataWinToSql(ddate) & "'")
    STB.Panels(4).Text = dat(1, 1)
    Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `otprvid` = 0")
     STB.Panels(6).Text = dat(1, 1)
End Sub

Private Sub mnuDir_Click()
If acl = "G" Or acl = "D" Then
frmdirectivi.Show vbModal, Me
End If
End Sub

Private Sub Timer2_Timer()
Static EventCount As Long
EventCount = EventCount + 1
If EventCount = 30 Then ' 30 это в секундах
Call info_panel
Call nom_prikaz
EventCount = 0
End If
End Sub


Private Sub FORm_LostFocus()
    Combo1.Enabled = False
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    
   
    cmdDiscon_Click
    
    End
End Sub




Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Combo1.Enabled = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Combo1.Enabled = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Combo1.Enabled = True
End Sub



Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Combo1.Enabled = False
End Sub

Private Sub gen01_Click()
frmPred.Show vbModal, Me

End Sub

Private Sub genalltoof_Click()
    Call Cnv.To_Listv
End Sub

Private Sub info_db_1_Click()
Call priziv_info
End Sub

Private Sub godr_Click()
Call Cnv.to_godr
End Sub

Private Sub mnu_log_Click()
If acl = "G" Then frmLog.Show vbModal, Me
End Sub


Private Sub mnu_nar_Click()
On Error Resume Next
If acl = "G" Or acl = "D" Then
    Call mysql.query("CHECK TABLE `naryad_srezki_" & nowBase & "`")
        If dat(3, 1) = "errOR" Then
            MsgBox "Извините эта функция доступна только с призыва `Осень 2007`", vbInformation, "Информация"
        Exit Sub
        Else
    frmNaryad.Show vbModal, Me
    End If
End If
End Sub

Private Sub mnu_stat_date_otp_Click()
Call Cnv.stat_otpravka
End Sub

Private Sub mnu_pr_old_Click()
frmdelprnik.Show vbModal, Me
End Sub
Private Sub mnu_service_spiski_ot_komp_Click()
If acl = "G" Then frmspiskikomp.Show vbModal, Me
 
End Sub
Private Sub mnu_service_vipa_d_Click()
'Call Cnv.vipa("0")
End Sub
Private Sub mnu_service_vipa_p_Click()
'Call Cnv.vipa("1")
End Sub
Private Sub mnu_settings_print_Click()
If acl = "G" Then frmset_print.Show vbModal, Me
End Sub

Private Sub mnu_stat_obraz_9k_fam_Click()
Call Cnv.obraz_fam(5)
End Sub

Private Sub mnu_stat_obraz_9k_vk_Click()
Call Cnv.obraz_vk(5)
End Sub

Private Sub mnu_stat_obraz_nsr_fam_Click()
Call Cnv.obraz_fam(4)
End Sub

Private Sub mnu_stat_obraz_nsr_vk_Click()
Call Cnv.obraz_vk(4)
End Sub

Private Sub mnu_stat_obraz_nvsh_fam_Click()
Call Cnv.obraz_fam(1)
End Sub

Private Sub mnu_stat_obraz_nvsh_vk_Click()
Call Cnv.obraz_vk(1)
End Sub

Private Sub mnu_stat_obraz_sr_fam_Click()
Call Cnv.obraz_fam(3)
End Sub

Private Sub mnu_stat_obraz_sr_vk_Click()
Call Cnv.obraz_vk(3)
End Sub

Private Sub mnu_stat_obraz_srspec_fam_Click()
Call Cnv.obraz_fam(2)
End Sub

Private Sub mnu_stat_obraz_srspec_vk_Click()
Call Cnv.obraz_vk(2)
End Sub

Private Sub mnu_stat_obraz_vsh_fam_Click()
Call Cnv.obraz_fam(0)
End Sub

Private Sub mnu_stat_obraz_vsh_vk_Click()
Call Cnv.obraz_vk(0)
End Sub

Private Sub mnu_system_settings_Click()
If acl = "G" Then frmSetting.Show vbModal, Me
End Sub

Private Sub mnu_system_users_Click()
If acl = "G" Then frmusers.Show vbModal, Me
End Sub

Public Sub mnu_table_check_Click()
If acl = "G" Then
    On Error Resume Next
    Dim x As Long
    Call sorting_blok("otpravkaid", "otpravka_" & nowBase & "")
    For x = 1 To UBound(outm())
        Call mysql.query("SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid='" & outm(x) & "'")
        Call mysql.query("UPDATE otpravka_" & nowBase & " set kolvo='" & dat(1, 1) & "' WHERE otpravkaid='" & outm(x) & "'")
    Next x
End If
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub mnuListsVK_Click()
    frmPrintListsVK.Show vbModal, Me
  
End Sub
Private Sub mnuRepORt_Click()
    frmReport.Show vbModal, Me
    
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show vbModal, Me
      End Sub

Private Sub mnuSettings_Click()
    frmSetting.Show vbModal, Me
End Sub

Private Sub optimiz_Click()
On Error Resume Next
If acl = "G" Then
Dim x As Long
Dim base As String
progress.Show
progress.Caption = "Оптимизация таблиц"
  Call mysql.query("show tables")
Dim datt() As String
datt() = dat()

'progress.ProgressBar1.Max = st
  For x = 1 To st
    base = datt(1, x)
    Call mysql.query("OPTIMIZE TABLE `" & base & "`")
  progress.ProgressBar1.Max = x
  Next x
Unload progress
MsgBox "Оптимизация таблиц завершена", vbInformation, "Оптимизация таблиц"
End If
End Sub

Private Sub repORt_to_date_Click()
frmFullReport.Show vbModal, Me
End Sub

Private Sub STB_PanelClick(ByVal Panel As MSComctlLib.Panel)
Call info_panel
End Sub

Private Sub mnu_stat_rodv_Click()
Call Cnv.rod
End Sub
Private Sub ip_upd()
On Error GoTo a
Dim objHTTP As Object
Set objHTTP = CreateObject("Microsoft.XMLHTTP")
Call objHTTP.Open("GET", "http://www.sun.autovv.ru/ip.php", False)
Call objHTTP.Send
'Call MsgBox(objHTTP.ResponseText)
a:
End Sub
