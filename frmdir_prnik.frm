VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmdir_prnik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Добавление/Редактирование директивщика"
   ClientHeight    =   8055
   ClientLeft      =   450
   ClientTop       =   420
   ClientWidth     =   6720
   Icon            =   "frmdir_prnik.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Адм. блокировка"
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Сохранить"
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   7440
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Основная информация"
      TabPicture(0)   =   "frmdir_prnik.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtFam"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtOtch"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtYear"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPrim"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lstvk"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lstdirs"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Блокировка"
      TabPicture(1)   =   "frmdir_prnik.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lfam"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lname"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lOtch"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lVk"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label18"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lYear"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label20"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label21"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lDataosp"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label23"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lkom"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label25"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lcom_date"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lstypk"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmdBlock"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.ComboBox lstdirs 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   -74760
         TabIndex        =   39
         Top             =   6120
         Width           =   5895
      End
      Begin VB.CommandButton cmdBlock 
         Caption         =   "Блокировка"
         Height          =   375
         Left            =   3720
         TabIndex        =   38
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox lstypk 
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
         Height          =   345
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox lstvk 
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
         Height          =   345
         Left            =   -71400
         TabIndex        =   17
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtPrim 
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
         Height          =   1215
         Left            =   -72720
         TabIndex        =   13
         Top             =   4200
         Width           =   4095
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
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
         Height          =   400
         Left            =   -74760
         TabIndex        =   11
         Top             =   4200
         Width           =   1200
      End
      Begin VB.TextBox txtOtch 
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
         Height          =   400
         Left            =   -74760
         TabIndex        =   8
         Top             =   2880
         Width           =   2755
      End
      Begin VB.TextBox txtName 
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
         Height          =   400
         Left            =   -71400
         TabIndex        =   6
         Top             =   1800
         Width           =   2755
      End
      Begin VB.TextBox txtFam 
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
         Height          =   400
         Left            =   -74760
         TabIndex        =   5
         Top             =   1800
         Width           =   2755
      End
      Begin VB.Label lcom_date 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   36
         Top             =   6600
         Width           =   2775
      End
      Begin VB.Label Label25 
         Caption         =   "Команда и дата"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   35
         Top             =   6600
         Width           =   2775
      End
      Begin VB.Label lkom 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   34
         Top             =   6000
         Width           =   2775
      End
      Begin VB.Label Label23 
         Caption         =   "В команде"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   33
         Top             =   6000
         Width           =   2775
      End
      Begin VB.Label lDataosp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   32
         Top             =   5400
         Width           =   2775
      End
      Begin VB.Label Label21 
         Caption         =   "Дата привоза"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   31
         Top             =   5400
         Width           =   2775
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Информация"
         BeginProperty Font 
            Name            =   "Arial CYR"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1800
         Width           =   5895
      End
      Begin VB.Label lYear 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   29
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Год рождения"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   28
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Label lVk 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   27
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Военкомат"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   26
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Label lOtch 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   25
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label Label14 
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label lname 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   23
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label lfam 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
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
         Left            =   3360
         TabIndex        =   21
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   20
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "УПК"
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
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Директива"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -74760
         TabIndex        =   14
         Top             =   5640
         Width           =   2755
      End
      Begin VB.Label Label8 
         Caption         =   "Примечание"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72720
         TabIndex        =   12
         Top             =   3720
         Width           =   2760
      End
      Begin VB.Label Label7 
         Caption         =   "Год рождения"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74760
         TabIndex        =   10
         Top             =   3720
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Военкомат"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   9
         Top             =   2520
         Width           =   2760
      End
      Begin VB.Label Label5 
         Caption         =   "Отчество"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74760
         TabIndex        =   7
         Top             =   2520
         Width           =   2760
      End
      Begin VB.Label Label4 
         Caption         =   "Имя"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71400
         TabIndex        =   4
         Top             =   1440
         Width           =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "Фамилия"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -74760
         TabIndex        =   3
         Top             =   1440
         Width           =   2755
      End
      Begin VB.Label lID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -71400
         TabIndex        =   2
         Top             =   840
         Width           =   2755
      End
      Begin VB.Label Label1 
         Caption         =   "ID директивщика"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -74760
         TabIndex        =   1
         Top             =   840
         Width           =   2755
      End
   End
End
Attribute VB_Name = "frmdir_prnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public type_load As Integer
Dim id_p As Long
Dim block As Boolean

Private Sub cmdAct_Click()
On Error GoTo ErrH
Dim idd As Long
Dim sql_in As String
Call mysql.query("select `id` from directivi_" & nowBase & " where `dir`='" & lstdirs & "'")
    idd = DAT(1, 1)
If type_load = "0" Then
    If Len(txtFam.Text) > 2 And Len(txtname.Text) > 1 And Len(txtOtch.Text) > 2 And Len(lstvk.Text) > 5 And Len(txtYear.Text) = 4 And Len(lstdirs) > 5 Then
        Call mysql.query("INSERT INTO `directivi_p_" & nowBase & "` (`id`,`fam`,`name`,`otch`,`vk`,`year_b`,`did`,`prim`) values ('" & lID & "','" & txtFam & "','" & txtname & "','" & txtOtch & "','" & lstvk.Text & "','" & txtYear & "','" & idd & "','" & txtPrim & "')")
        MsgBox "Директивщик удачно был добавлен в Базу", vbInformation, "Добавлен"
    Else
        MsgBox "Вы ввели не всю информацию!!!", vbCritical, "Добавление"
    End If
Else
    Call mysql.query("SELECT `id`,`pid` from directivi_p_" & nowBase & " where `id`='" & lID & "'")
    datt() = DAT()
    Call mysql.query("DELETE from directivi_p_" & nowBase & " WHERE `id`='" & lID & "'")
    If block = True Then
        Call mysql.query("INSERT INTO directivi_p_" & nowBase & " (`id`,`fam`,`name`,`otch`,`vk`,`year_b`,`did`,`pid`,`prim`) VALUES ('" & datt(1, 1) & "','" & txtFam & "','" & txtname & "','" & txtOtch & "','" & lstvk & "','" & txtYear & "','" & idd & "','" & datt(2, 1) & "','" & txtPrim & "')")
    Else
        Call mysql.query("INSERT INTO directivi_p_" & nowBase & " (`id`,`fam`,`name`,`otch`,`vk`,`year_b`,`did`,`pid`,`prim`) VALUES ('" & datt(1, 1) & "','" & txtFam & "','" & txtname & "','" & txtOtch & "','" & lstvk & "','" & txtYear & "','" & idd & "','','" & txtPrim & "')")
    End If
    Call load_info(lID)
End If
Unload Me
Exit Sub
ErrH:
MsgBox "ERROR", vbCritical, "ERROR"
End Sub

Private Sub cmdBlock_Click()

If block = True Then
    Call unblock_p(lstypk.Text)
    block = False
    cmdblock.Caption = "Блокировать"
    Call Form_Load
Else
    Call block_p(lstypk.Text)
    block = True
    cmdblock.Caption = "Разблокировать"
    load_info (frmdir.lstprnik.SelectedItem.Text)
    Call Form_Load
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub CommAND3_Click()
On Error Resume Next
Dim PID As Long
PID = InputBox("Введите номер УПК которому соответствует данный директивщик", "Административная блокировка директивщика")
Call mysql.query("SELECT `fam`,`name`,`otch`,`txtvk`,`datar`,`dataosp` from prnik_" & nowBase & " where `idprnik`='" & PID & "'")
If val(st) > 0 Then
    If MsgBox("Короткая информация о призывнике с УПК '" & PID & "' :" & Chr(10) & "Ф.И.О. :" & DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & Chr(10) & "ВК :" & DAT(4, 1) & Chr(10) & "Год рождения :" & Left(DAT(5, 1), 4) & Chr(10) & "Приехал на ОСП :" & CnvDataSqLToWin(DAT(6, 1)) & Chr(10) & "Блокировать?", vbYesNo + vbInformation, "Блокировка") = vbYes Then
        block_p (PID)
        MsgBox "Призывник с УПК " & PID & " заблокирован!", vbInformation + vbOKOnly, "Блокировка"
    End If
Else
    MsgBox "Извините, но данного УПК в базе нету!!!", vbCritical + vbOKOnly, "Блокировка"
End If
Call Form_Load
End Sub

Private Sub block_p(PID As Long)
Call mysql.query("UPDATE `directivi_p_" & nowBase & "` set `pid`='" & PID & "' where `id`='" & lID & "'  ")
Call mysql.query("update `prnik_" & nowBase & "` set `lock`='3' where `idprnik`='" & PID & "'")
Call mysql.query("update `prnik_" & nowBase & "` set `lprim`='Директивщик: директива " & lstdirs.Text & " ' where `idprnik`='" & PID & "'")
End Sub
Private Sub unblock_p(PID As Long)
Call mysql.query("UPDATE `directivi_p_" & nowBase & "` set `pid`='' where `id`='" & lID & "'")
Call mysql.query("update `prnik_" & nowBase & "` set `lock`='' where `idprnik`='" & PID & "'")
Call mysql.query("update `prnik_" & nowBase & "` set `lprim`='' where `idprnik`='" & PID & "'")
End Sub

Private Sub Form_Load()
On Error Resume Next
block = False
lstypk.Enabled = True
If type_load = "0" Then Call new_prnik
If type_load = "1" Then Call load_info(frmdir.lstprnik.SelectedItem.Text)
Call mysql.query("select `dir` from `directivi_" & nowBase & "` ORDER BY `dir` ASC")
For x = 1 To st
    lstdirs.AddItem (DAT(1, x))
Next x
Call Reg_VK_List
For x = 0 To UBound(nVK())
    lstvk.AddItem (nVK(x))
Next x
End Sub
Private Sub load_info(ID As Long)
On Error Resume Next
cmdAct.Caption = "Сохранить"
id_p = frmdir.lstprnik.SelectedItem.Text
lID = id_p
lstdirs.Text = frmdir.lstdirs.SelectedItem.SubItems(1)
Call mysql.query("SELECT `fam`,`name`,`otch`,`vk`,`year_b`,`pid`,`prim` from directivi_p_" & nowBase & " where `id`='" & id_p & "'")

txtFam = DAT(1, 1)
txtname = DAT(2, 1)
txtOtch = DAT(3, 1)
lstvk.Text = DAT(4, 1)
txtYear = DAT(5, 1)
txtPrim = DAT(7, 1)
If val(DAT(6, 1)) > 0 Then
    block = True
    lstypk = DAT(6, 1)
    lstypk.Enabled = False
    cmdblock.Caption = "Разблокировать"
    Call mysql.query("SELECT `fam`,`name`,`otch`,`txtvk`,`datar`,`dataosp`,`otprvid` from prnik_" & nowBase & " where `idprnik`='" & lstypk.Text & "'")
    lfam = DAT(1, 1)
    lname = DAT(2, 1)
    lOtch = DAT(3, 1)
    lstvk = DAT(4, 1)
    lYear = Left(DAT(5, 1), 4)
    lDataosp = CnvDataSqLToWin(DAT(6, 1))
    If val(DAT(7, 1)) > 0 Then
        lkom = "ДА"
        Call mysql.query("SELECT naryad_" & nowBase & ".`oblkom`,otpravka_" & nowBase & ".`data` from naryad_" & nowBase & ", otpravka_" & nowBase & " where otpravka_" & nowBase & ".`otpravkaid` = '" & DAT(7, 1) & "' and naryad_" & nowBase & ".`narid`=otpravka_" & nowBase & ".`narid`")
        lcom_date = DAT(1, 1) & " от " & DAT(2, 1)
    End If
    
End If
 Call mysql.query("select `idprnik` from `prnik_" & nowBase & "` where `fam` like '" & Left$(frmdir.lstprnik.SelectedItem.SubItems(1), 4) & "%' and `txtvk` = '" & frmdir.lstprnik.SelectedItem.SubItems(4) & "'")
    lstypk.Clear
    If val(st) > 0 Then
        For x = 1 To st
            lstypk.AddItem (DAT(1, x))
        Next x
        lstypk.ListIndex = 0
    End If

End Sub

Private Sub new_prnik()
On Error Resume Next
cmdAct.Caption = "Добавить"
SSTab1.TabEnabled(1) = False
id_p = get_new_id
lID = id_p
lstdirs = frmdir.lstdirs.SelectedItem.SubItems(1)
End Sub
Private Function get_new_id() As Long
Call mysql.query("select max(id) from directivi_p_" & nowBase)
get_new_id = val(DAT(1, 1)) + 1
End Function

Private Sub Form_Unload(Cancel As Integer)
frmdir.sum_dir
frmdir.lstdirs_Click
frmdir.view_nonblock
End Sub

Private Sub lstypk_Click()
Call mysql.query("SELECT `fam`,`name`,`otch`,`txtvk`,`datar`,`dataosp`,`otprvid` from prnik_" & nowBase & " where `idprnik`='" & lstypk.Text & "'")
    lfam = DAT(1, 1)
    lname = DAT(2, 1)
    lOtch = DAT(3, 1)
    lVk = DAT(4, 1)
    lYear = Left(DAT(5, 1), 4)
    lDataosp = CnvDataSqLToWin(DAT(6, 1))
    If val(DAT(7, 1)) > 0 Then
        lkom = "ДА"
        Call mysql.query("SELECT naryad_" & nowBase & ".`oblkom`,otpravka_" & nowBase & ".`data` from naryad_" & nowBase & ", otpravka_" & nowBase & " where otpravka_" & nowBase & ".`otpravkaid` = '" & DAT(7, 1) & "' and naryad_" & nowBase & ".`narid`=otpravka_" & nowBase & ".`narid`")
        lcom_date = "'" & DAT(1, 1) & "' от " & CnvDataSqLToWin(DAT(2, 1))
    End If
End Sub
