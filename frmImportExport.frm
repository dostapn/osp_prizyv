VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Экспорт"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImportExport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Начать"
      Default         =   -1  'True
      Height          =   360
      Left            =   5505
      TabIndex        =   12
      Top             =   5160
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Сохранить"
      Height          =   3870
      Left            =   150
      TabIndex        =   3
      Top             =   1215
      Width           =   6495
      Begin VB.CheckBox EXP_chDoEvents 
         Caption         =   "Построчное сохранение"
         Enabled         =   0   'False
         Height          =   255
         Left            =   255
         TabIndex        =   15
         Top             =   3450
         Value           =   1  'Отмечено
         Width           =   3945
      End
      Begin VB.CheckBox EXP_chData 
         Caption         =   "Только данные"
         Enabled         =   0   'False
         Height          =   225
         Left            =   615
         TabIndex        =   14
         Top             =   1125
         Value           =   1  'Отмечено
         Width           =   3975
      End
      Begin VB.CheckBox EXP_chStruct 
         Caption         =   "Только структуру"
         Enabled         =   0   'False
         Height          =   285
         Left            =   615
         TabIndex        =   13
         Top             =   735
         Width           =   4395
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Обзор..."
         Height          =   360
         Left            =   5250
         TabIndex        =   11
         Top             =   2985
         Width           =   1125
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Плоска
         BorderStyle     =   0  'Нет
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   585
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3075
         Width           =   4515
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Выбрать..."
         Height          =   360
         Left            =   5250
         TabIndex        =   8
         Top             =   2220
         Width           =   1125
      End
      Begin VB.TextBox txtSel 
         Appearance      =   0  'Плоска
         BorderStyle     =   0  'Нет
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   585
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2295
         Width           =   4530
      End
      Begin VB.OptionButton EXP_optSel 
         Caption         =   "Выбранная:"
         Height          =   225
         Left            =   255
         TabIndex        =   6
         Top             =   1905
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.CheckBox EXP_chAll 
         Caption         =   "Структуру и данные всех таблиц"
         Enabled         =   0   'False
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   405
         Width           =   4725
      End
      Begin VB.OptionButton EXP_optAllTab 
         Caption         =   "Все таблицы"
         Enabled         =   0   'False
         Height          =   300
         Left            =   255
         TabIndex        =   4
         Top             =   1485
         Width           =   2865
      End
      Begin VB.Label lblDir 
         BackStyle       =   0  'Прозрачно
         Caption         =   "В директорию:"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   2685
         Width           =   2700
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Непрозрачно
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   495
         Top             =   3030
         Width           =   4650
      End
      Begin VB.Shape ShapeUpk 
         BackStyle       =   1  'Непрозрачно
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   495
         Top             =   2250
         Width           =   4665
      End
   End
   Begin VB.PictureBox pic12 
      Align           =   1  'Привязать вверх
      BackColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   0
      Picture         =   "frmImportExport.frx":0CCA
      ScaleHeight     =   1050
      ScaleWidth      =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   6780
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Выберите необходимую категорию"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   2670
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Утилита для резервного сохнанения открытой базы данных"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   990
         TabIndex        =   1
         Top             =   270
         Width           =   5430
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   255
         Picture         =   "frmImportExport.frx":3188
         Top             =   285
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
