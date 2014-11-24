VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInfoPr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Информация о призывнике"
   ClientHeight    =   7650
   ClientLeft      =   5865
   ClientTop       =   3480
   ClientWidth     =   11490
   Icon            =   "frmInfoPr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Панель меню"
      Height          =   7455
      Left            =   9840
      TabIndex        =   56
      Top             =   120
      Width           =   1575
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   780
         Left            =   480
         Picture         =   "frmInfoPr.frx":08CA
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   69
         Top             =   6240
         Width           =   780
      End
      Begin VB.CommandButton cmd_stat 
         Caption         =   "Статистика"
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Печать"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmd_del 
         Caption         =   "Удалить"
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_go 
         Caption         =   "Перейти"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmd_fromKom 
         Caption         =   "Из команды"
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "Сохранить"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Отмена"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "ОК"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7605
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   13414
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "Общие"
      TabPicture(0)   =   "frmInfoPr.frx":1044
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Дополнительно"
      TabPicture(1)   =   "frmInfoPr.frx":1060
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Управление"
      TabPicture(2)   =   "frmInfoPr.frx":107C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Статистика"
      TabPicture(3)   =   "frmInfoPr.frx":1098
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "stat"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.ListView stat 
         Height          =   6735
         Left            =   -74760
         TabIndex        =   74
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial CYR"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Действие"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Куда/Откуда"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Дата"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Кто"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Общая информация"
         Height          =   6855
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   9495
         Begin VB.Frame Frame9 
            Caption         =   "Отправка в войска"
            Height          =   4335
            Left            =   4800
            TabIndex        =   12
            Top             =   2400
            Width           =   4575
            Begin VB.TextBox txtdir 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1920
               TabIndex        =   73
               Top             =   3960
               Width           =   2535
            End
            Begin VB.ComboBox vys 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               TabIndex        =   72
               Top             =   2160
               Width           =   2535
            End
            Begin VB.Label Label38 
               Caption         =   "Директива"
               Height          =   255
               Left            =   240
               TabIndex        =   68
               Top             =   3960
               Width           =   1095
            End
            Begin VB.Label txtforvch 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   1920
               TabIndex        =   55
               Top             =   3360
               Width           =   2535
            End
            Begin VB.Label txtforpunkt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   54
               Top             =   3000
               Width           =   2535
            End
            Begin VB.Label txtoblkom 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   53
               Top             =   2640
               Width           =   2535
            End
            Begin VB.Label txtokrug 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   52
               Top             =   1800
               Width           =   2535
            End
            Begin VB.Label Label27 
               Caption         =   "Для части"
               Height          =   255
               Left            =   240
               TabIndex        =   51
               Top             =   3360
               Width           =   1095
            End
            Begin VB.Label Label26 
               Caption         =   "Для пункта"
               Height          =   255
               Left            =   240
               TabIndex        =   50
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Label25 
               Caption         =   "Обл. команда"
               Height          =   255
               Left            =   240
               TabIndex        =   49
               Top             =   2640
               Width           =   1095
            End
            Begin VB.Label Label24 
               Caption         =   "ВУС"
               Height          =   255
               Left            =   240
               TabIndex        =   48
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label23 
               Caption         =   "Военный округ"
               Height          =   255
               Left            =   240
               TabIndex        =   47
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label txtvch 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   46
               Top             =   1440
               Width           =   2535
            End
            Begin VB.Label Label21 
               Caption         =   "Воинская часть"
               Height          =   255
               Left            =   240
               TabIndex        =   45
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label txtpunkt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   44
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label19 
               Caption         =   "Пункт дислокации"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label txtrodv 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   42
               Top             =   720
               Width           =   2535
            End
            Begin VB.Label Label17 
               Caption         =   "Род войск"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label txtotp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1920
               TabIndex        =   40
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label15 
               Caption         =   "Отправлен"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Остальная"
            Height          =   1935
            Left            =   4800
            TabIndex        =   11
            Top             =   360
            Width           =   4455
            Begin VB.CheckBox vod 
               Caption         =   "Check1"
               Height          =   255
               Left            =   1320
               TabIndex        =   88
               Top             =   1560
               Width           =   255
            End
            Begin VB.ComboBox sud 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1320
               TabIndex        =   37
               Top             =   960
               Width           =   2895
            End
            Begin VB.ComboBox sempol 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1320
               TabIndex        =   35
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label Label14 
               Caption         =   "Водитель"
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label13 
               Caption         =   "Судимость"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Семейное положение"
               Height          =   375
               Left            =   240
               TabIndex        =   34
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Общая"
            Height          =   6255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   4095
            Begin VB.TextBox txtvus_p 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   89
               Top             =   5760
               Width           =   2535
            End
            Begin VB.TextBox vb2 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1920
               TabIndex        =   33
               Top             =   5280
               Width           =   1935
            End
            Begin VB.TextBox vb1 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   32
               Top             =   5280
               Width           =   375
            End
            Begin VB.TextBox txtdatapr 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   30
               Top             =   4800
               Width           =   2535
            End
            Begin VB.ComboBox vklist 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1320
               TabIndex        =   28
               Top             =   4320
               Width           =   2535
            End
            Begin VB.ComboBox spec 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1440
               TabIndex        =   26
               Top             =   3600
               Width           =   2415
            End
            Begin VB.ComboBox obrazov 
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   1440
               TabIndex        =   25
               Top             =   3120
               Width           =   2415
            End
            Begin VB.TextBox txtdatar 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1800
               TabIndex        =   22
               Top             =   2400
               Width           =   2055
            End
            Begin VB.TextBox txtotch 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   20
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox txtname 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   19
               Top             =   1440
               Width           =   2535
            End
            Begin VB.TextBox txtfam 
               BackColor       =   &H00C0FFC0&
               Height          =   285
               Left            =   1320
               TabIndex        =   18
               Top             =   960
               Width           =   2535
            End
            Begin VB.Label Label22 
               Caption         =   "Личный номер"
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   5760
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Военный билет"
               Height          =   375
               Left            =   120
               TabIndex        =   31
               Top             =   5280
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Дата призыва"
               Height          =   495
               Left            =   120
               TabIndex        =   29
               Top             =   4800
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Военкомат"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   4320
               Width           =   855
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4080
               Y1              =   4200
               Y2              =   4200
            End
            Begin VB.Label Label8 
               Caption         =   "Специальность"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   3600
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Образование"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4080
               Y1              =   2880
               Y2              =   2880
            End
            Begin VB.Label Label6 
               Caption         =   "Дата рождения"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   2400
               Width           =   1695
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4080
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Label txtid 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1200
               TabIndex        =   17
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label Label4 
               Caption         =   "Отчество"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Имя"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Фамилия"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "УПК"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Управление"
         Height          =   6855
         Left            =   -74760
         TabIndex        =   2
         Top             =   600
         Width           =   9495
         Begin VB.TextBox txtlockpr 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2280
            MultiLine       =   -1  'True
            TabIndex        =   70
            Top             =   1200
            Width           =   6735
         End
         Begin VB.ComboBox lstblock 
            BackColor       =   &H00C0FFC0&
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label40 
            Caption         =   "Причина блокировки, кем и куда"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            TabIndex        =   71
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Тип блокировки"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Дополнительная информация (разработка)"
         Height          =   6855
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   9495
         Begin VB.Frame Frame11 
            Caption         =   "Медицина"
            Height          =   2055
            Left            =   4920
            TabIndex        =   75
            Top             =   360
            Width           =   3975
            Begin VB.Label lmedst 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   2280
               TabIndex        =   79
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label16 
               Caption         =   "Медицинская статья"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   240
               TabIndex        =   78
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label lmed 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   400
               Left            =   2280
               TabIndex        =   77
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Степень годности"
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
               Left            =   240
               TabIndex        =   76
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Дополнительная информация"
            Height          =   3975
            Left            =   5040
            TabIndex        =   65
            Top             =   2640
            Width           =   3855
            Begin VB.Label ldop 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1935
               Left            =   120
               TabIndex        =   87
               Top             =   1800
               Width           =   3615
            End
            Begin VB.Label Label20 
               Caption         =   "Прочая информация"
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
               Left            =   240
               TabIndex        =   86
               Top             =   1440
               Width           =   2295
            End
            Begin VB.Label lvus 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   120
               TabIndex        =   81
               Top             =   840
               Width           =   3615
            End
            Begin VB.Label Label18 
               Caption         =   "Первоначальный ВУС"
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
               Left            =   240
               TabIndex        =   80
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Домашний адресс"
            Height          =   3375
            Left            =   240
            TabIndex        =   6
            Top             =   3360
            Width           =   4095
            Begin VB.Label ltel 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               TabIndex        =   85
               Top             =   2880
               Width           =   3800
            End
            Begin VB.Label ladr 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1695
               Left            =   120
               TabIndex        =   84
               Top             =   720
               Width           =   3800
            End
            Begin VB.Label Label34 
               Caption         =   "Телефон"
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
               Left            =   240
               TabIndex        =   8
               Top             =   2520
               Width           =   1215
            End
            Begin VB.Label Label33 
               Caption         =   "Адресс"
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
               Left            =   240
               TabIndex        =   7
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Родители"
            Height          =   2895
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   4095
            Begin VB.Label lmatj 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   700
               Left            =   120
               TabIndex        =   83
               Top             =   1920
               Width           =   3800
            End
            Begin VB.Label lotec 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   700
               Left            =   120
               TabIndex        =   82
               Top             =   720
               Width           =   3800
            End
            Begin VB.Label Label32 
               Caption         =   "Мать"
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
               Left            =   240
               TabIndex        =   5
               Top             =   1560
               Width           =   495
            End
            Begin VB.Label Label30 
               Caption         =   "Отец"
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
               Left            =   240
               TabIndex        =   4
               Top             =   360
               Width           =   495
            End
         End
      End
   End
End
Attribute VB_Name = "frmInfoPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim savea As Boolean
Dim P_txtfam As String
Dim P_txtname As String
Dim P_txtotch As String
Dim P_txtdatar As String
Dim P_txtdatapr As String
Dim P_vb1 As String
Dim P_vb2 As String
Dim P_obrazov As String
Dim P_spec As String
Dim P_vklist As String
Dim P_sempol As String
Dim P_sud As String
Dim P_vod As Long
Dim P_vus As String
Dim P_dir As String
Dim P_txtlockpr As String
Dim P_lock_u As Integer
Dim P_txtvus_p As String
Dim idp As Integer
Dim lock_u As Integer
Dim dir As String
Dim vus_now As String


Private Sub cmd_ok_Click()
On Error Resume Next
If acl = "G" Or acl = "D" Then
If txtfam = P_txtfam And txtname = P_txtname And txtotch = P_txtotch And txtdatar = P_txtdatar And txtvus_p = P_txtvus_p And txtdatapr = P_txtdatapr And vb1 = P_vb1 And vb2 = P_vb2 And obrazov = P_obrazov And spec = P_spec And vklist = P_vklist And sempol = P_sempol And sud = P_sud And vod.VAlue = P_vod And vys = P_vus And txtDir = P_dir And txtlockpr = P_txtlockpr And lstblock.ListIndex = P_lock_u Then
            If frmCommand.Visible = True Then frmCommand.commANDs_load
            If frmSearch.Visible = True Then frmSearch.cmdSearch_Click
    Unload Me
Else
    If Not acl = "s" Then
        If MsgBox("Вы изминили данные о призывнике! Сохранить эти изменения?", vbYesNo + vbInformation, "Сохранение призывника") = vbYes Then
            Call cmd_save_Click
            
            If frmCommand.Visible = True Then frmCommand.commANDs_load
            If frmSearch.Visible = True Then frmSearch.cmdSearch_Click
            

            Unload Me
        End If
    Else
      MsgBox "Извините Вашему пользователю запрещен доступ на изменение данных о призывниках!", vbInformation, "Сохранение призывника"
    End If
End If
End If
'If frmSearch.Visible = True Then frmSearch.txtpole.SetFocus
End Sub
Private Sub cmd_cancel_Click()
Unload Me
End Sub
Private Sub cmd_del_Click()
If acl = "G" Or acl = "D" Then
frmdelprnik.Show vbModal, Me
End If
End Sub

Private Sub cmd_fromKom_Click()
On Error Resume Next
cUpk = CLng(txtid.Caption)
If MsgBox("Вы уверены что хотите удалить этого призывника из команды", vbOKCancel, "Удаления призывника из команды") = vbOK Then


Call mysql.query("SELECT `lock` FROM prnik_" & nowBase & " WHERE idprnik = '" & cUpk & "'")
    If CLng(DAT(1, 1)) = "3" Then
            If acl = "G" Then
               GoTo del_fr
            Else
                MsgBox "Извините Вашему пользователю запрещен доступ!", vbInformation, "Доступ запрещен!": Exit Sub
            End If
        End If
    
    If CLng(DAT(1, 1)) = "2" Then MsgBox "Этот призывник нахидся в ушедшой команде", vbInformation, "Доступ запрещен!": Exit Sub

GoTo del_fr
    
del_fr:
            Call mysql.query("SELECT count(idprnik) FROM prnik_" & nowBase & " WHERE otprvid=(select `otprvid` from prnik_" & nowBase & " where idprnik='" & cUpk & "')")
            Call mysql.query("UPDATE otpravka_" & nowBase & " set kolvo='" & DAT(1, 1) - 1 & "' WHERE otpravkaid=(select `otprvid` from prnik_" & nowBase & " where idprnik='" & cUpk & "')")
            Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `otprvid` = '0' WHERE `idprnik` = '" & cUpk & "'")
            Call mysql.query("UPDATE `prnik_" & nowBase & "` SET `vus` = '' WHERE `idprnik` = '" & cUpk & "'")
            Dim komm As String
            Call refresh_info
End If
    
End Sub

Private Sub cmd_go_Click()
'On ErrOR Resume Next
'Dim datetmp() As String
'dateotp = CnvDataWinToSql(Label16.Caption)
'datetmp() = Split(CnvDataSqLToWin(dateotp), ".")

'p_com_id = 0

'frmCommAND.Show vbModal, Me

'frmCommAND.listcommANDs.ListItems.Clear
'frmCommAND.listprnik.ListItems.Clear
'frmCommAND.listcommANDs.SELECTedItem.SELECTed = False
'frmCommAND.listcommANDs.FindItem(Label31.Caption).SELECTed = True
'frmCommAND.listcommANDs.FindItem(Label31.Caption).EnsureVisible = True
'frmCommAND.commANDs_load
'frmCommAND.prnik_load
'frmCommAND.listprnik.FindItem(Label5.Caption).SELECTed = True
'frmCommAND.listprnik.FindItem(Label5.Caption).EnsureVisible = True
'
End Sub



Private Sub cmd_save_Click()
On Error Resume Next
If acl = "G" Or acl = "D" Then
Err = vbNull
Dim strIn As String

Call mysql.query("update prnik_" & nowBase & " set `txtvk` = '" & vklist.Text & "',`fam` = '" & txtfam.Text & "',`name` = '" & txtname.Text & "',`otch` = '" & txtotch.Text & "',`datar` = '" & CnvDataWinToSql(txtdatar.Text) & "',`txtobraz` = '" & obrazov.Text & "',`txtspec` = '" & spec.Text & "',`txtsem` = '" & sempol.Text & "',`txtsud` = '" & sud.Text & "',`servb` = '" & vb1.Text & "',`nomvb` = '" & vb2.Text & "',`dataosp` = '" & CnvDataWinToSql(txtdatapr.Text) & "',`vod` = " & vod.VAlue & ",`lock` = '" & lstblock.ListIndex & "',`lprim` = '" & txtlockpr.Text & "',`vus_p` = '" & txtvus_p.Text & "' WHERE `idprnik`='" & txtid.Caption & "'")
MsgBox "Информация о призывнике " & txtfam & " была удачно сохранена", vbInformation, "Сохранение изменений"
Call log_sql("0", "2", expupk, "Данные")
Call frmSearch.cmdSearch_Click
Call refresh_info

End If

End Sub



Private Sub cmd_print_Click()
On Error Resume Next
''Screen.MousePointer = vbHourglass
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "info.xls"
    If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
With oWs
          
                       
            .cells(1, 2) = txtid.Caption
            .cells(3, 2) = txtfam
            .cells(4, 2) = txtname
            .cells(5, 2) = txtotch
            .cells(6, 2) = txtdatar
            .cells(8, 2) = obrazov
            .cells(9, 2) = spec
            .cells(11, 2) = vklist
            .cells(12, 2) = txtdatapr
            .cells(13, 2) = vb1 & " " & vb2
            .cells(15, 2) = sempol
            .cells(16, 2) = sud
            If vod.VAlue = 1 Then .cells(17, 2) = "ДА" Else .cells(17, 2) = "НЕТ"
            .cells(19, 2) = txtotp
            .cells(20, 2) = txtrodv
            .cells(21, 2) = txtpunkt
            .cells(22, 2) = txtvch
            .cells(23, 2) = txtokrug
            .cells(24, 2) = vys
            .cells(25, 2) = txtoblkom
            .cells(26, 2) = txtforpunkt
            .cells(27, 2) = txtforvch
            .cells(28, 2) = txtDir
        
            
End With
End Sub

Private Sub CommAND9_Click()

End Sub

Private Sub cmd_stat_Click()
frmstatprnik.Show vbModal, Me
End Sub

Private Sub Form_Load()
'Call mysql.free_result
On Error Resume Next
savea = False
Call Reg_VK_List
For x = 1 To UBound(nVK())
    vklist.AddItem (nVK(x))
Next x

Call obraz_in
For x = 1 To UBound(obraz())
obrazov.AddItem (obraz(x))
Next x

sempol.AddItem ("Холост")
sempol.AddItem ("Женат")
sempol.AddItem ("Разведен")
sempol.AddItem ("Вдовец")

sud.AddItem ("Не судим")
sud.AddItem ("Судим")
sud.AddItem ("Приводы")


'Безопасность
If acl = "s" Then
    cmd_save.Enabled = False
    cmd_fromKom.Enabled = False
    cmd_del.Enabled = False
End If

Call refresh_info
Call stat_prnik

End Sub
Private Sub stat_prnik()
On Error Resume Next
Dim argw, time, who As String
Dim data As Date
Dim x, c As Long
Dim datt() As String
Call log_types
Call mysql.query("SELECT  `type`,`act`,`argw`,`data`,`who` FROM logs_" & nowBase & " WHERE type='0' AND to_id='" & expupk & "' ORder by data,time ASC")
datt() = DAT()
stat.ListItems.Clear
For x = 1 To st
    datt(1, x) = log_type_act(datt(1, x))
    datt(2, x) = log_act(datt(2, x))
    datt(4, x) = CnvDataSqLToWin(datt(4, x))
    datt(5, x) = get_fio(DAT(5, x))
    Set LF = stat.ListItems.add()
    For c = 2 To 5
        LF.SubItems(c - 1) = datt(c, x)
    Next c
Next x
Call ReSizeColumnHeaders(stat)
End Sub
Private Sub clear_info()
txtfam = vbNullString
txtname = vbNullString
txtotch = vbNullString
txtdatar = vbNullString
txtdatapr = vbNullString
vb1 = vbNullString
vb2 = vbNullString
vod_y = False
vod_n = False
txtotp = vbNullString
txtrodv = vbNullString
txtpunkt = vbNullString
txtvch = vbNullString
txtokrug = vbNullString
txtoblkom = vbNullString
txtforpunkt = vbNullString
txtforvch = vbNullString
vys = vbNullString
txtDir.Text = vbNullString
txtmatj = vbNullString
txtotec = vbNullString
txtadr = vbNullString
txttel = vbNullString

End Sub

Private Sub refresh_info()
On Error Resume Next

Call clear_info

lupk = expupk
Call mysql.query("SELECT `fam`,`name`,otch,datar,dataosp,servb,nomvb,txtobraz, txtspec, txtvk, txtsem, txtsud, `vod`, `otprvid` ,`dir`,`lock`,`lprim` , `vus_p` FROM prnik_" & nowBase & " WHERE idprnik='" & lupk & "'")
P_txtfam = DAT(1, 1)
P_txtname = DAT(2, 1)
P_txtotch = DAT(3, 1)
P_txtdatar = CnvDataSqLToWin(DAT(4, 1))
P_txtdatapr = CnvDataSqLToWin(DAT(5, 1))
P_vb1 = DAT(6, 1)
P_vb2 = DAT(7, 1)
P_obrazov = DAT(8, 1)
P_spec = DAT(9, 1)
P_vklist = DAT(10, 1)
P_sempol = DAT(11, 1)
P_sud = DAT(12, 1)
P_txtvus_p = DAT(18, 1)
P_vod = DAT(13, 1)
''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''
idp = DAT(14, 1)
''''''''''''
P_lock_u = DAT(16, 1)
P_dir = DAT(15, 1)
P_txtlockpr = DAT(17, 1)
lock_u = P_lock_u
dir = P_dir
txtlockpr = P_txtlockpr
''''''''''''''''''''''''
''''Заполнение полей''''
''''''''''''''''''''''''
If DAT(13, 1) = "1" Then vod.VAlue = 1 Else vod.VAlue = 0
txtid = lupk
txtfam = P_txtfam
txtname = P_txtname
txtotch = P_txtotch
txtdatar = P_txtdatar
txtdatapr = P_txtdatapr
txtvus_p = P_txtvus_p
vb1 = P_vb1
vb2 = P_vb2
obrazov = P_obrazov
spec = P_spec
vklist = P_vklist
sempol = P_sempol
sud = P_sud
Call mysql.query("SELECT data, naryad_" & nowBase & ".rodv, naryad_" & nowBase & ".punkt, naryad_" & nowBase & ".vch, naryad_" & nowBase & ".okr, naryad_" & nowBase & ".oblkom,fORpunkt,fORchast FROM otpravka_" & nowBase & ",naryad_" & nowBase & " WHERE otpravka_" & nowBase & ".otpravkaid='" & idp & "' AND otpravka_" & nowBase & ".narid=naryad_" & nowBase & ".narid")
If st > "0" Then
txtotp.Caption = CnvDataSqLToWin(DAT(1, 1))
txtrodv.Caption = DAT(2, 1)
txtpunkt.Caption = DAT(3, 1)
txtvch.Caption = DAT(4, 1)
txtokrug.Caption = DAT(5, 1)
txtoblkom.Caption = DAT(6, 1)
txtforpunkt.Caption = DAT(7, 1)
txtforvch.Caption = DAT(8, 1)
txtDir.Text = dir
End If
If lock_u = "3" Then
    If acl = "G" Then
        Resume
    Else
        lstblock.Enabled = False
    End If
End If
    Call mysql.query("SELECT vus FROM prnik_" & nowBase & " WHERE idprnik='" & lupk & "'")
    P_vus = DAT(1, 1)
    vus_now = P_vus
    
    Call mysql.query("SELECT okrkom FROM naryad_" & nowBase & " WHERE oblkom='" & txtoblkom.Caption & "'")
    If st > "0" Then
        Call mysql.query("SELECT vrp FROM naryad_" & nowBase & " WHERE okrkom='" & DAT(1, 1) & "'")
            For x = 1 To st
                vys.AddItem (DAT(1, x))
            Next x
            vys.Text = vus_now
    End If
lstblock.AddItem ("Нету")
lstblock.AddItem ("Обычная")
lstblock.AddItem ("Командная")
lstblock.AddItem ("Администраторская")
lstblock.ListIndex = lock_u

  Call mysql.query("CHECK TABLE `prnik_" & nowBase & "`")
            If DAT(3, 1) = "errOR" Then
            frmInfoPr.SSTab1.TabVisible(1) = False
            Else
            If st > 0 Then
                Call mysql.query("SELECT `vus`,`matj`,`otec`,`adr`,`dop`,`tel`,`med`,`medst` FROM prnik_" & nowBase & " WHERE idprnik='" & lupk & "'")
                    lvus = DAT(1, 1)
                    lmatj = DAT(2, 1)
                    lotec = DAT(3, 1)
                    ladr = DAT(4, 1)
                    ldop = DAT(5, 1)
                    ltel = DAT(6, 1)
                    lmed = DAT(7, 1)
                    lmedst = DAT(8, 1)
            End If
            End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmCommand.Visible = True Then frmCommand.commANDs_load
            If frmSearch.Visible = True Then frmSearch.cmdSearch_Click
End Sub


Private Sub lstblock_Click()
If lstblock.ListIndex = "3" Then
    If Not acl = "G" Then
        MsgBox "Извините Вашему пользователю запрещен доступ!", vbInformation, "Доступ запрещен!"
        Exit Sub
    End If
End If
End Sub
