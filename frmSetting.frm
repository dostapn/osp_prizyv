VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Настройки"
   ClientHeight    =   5685
   ClientLeft      =   405
   ClientTop       =   435
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ок"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   360
      Left            =   7050
      TabIndex        =   2
      Top             =   5265
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   360
      Left            =   8145
      TabIndex        =   1
      Top             =   5265
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5085
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   10
      TabHeight       =   520
      WordWrap        =   0   'False
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Общие"
      TabPicture(0)   =   "frmSetting.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lnprik"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Calendar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "prodprik"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Change"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Настройки наряда"
      TabPicture(1)   =   "frmSetting.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "txtrodf"
      Tab(1).Control(3)=   "txtpodrodf"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Настройка ВИПА"
      TabPicture(2)   =   "frmSetting.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Change 
         Caption         =   "Изменить"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox prodprik 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtpodrodf 
         Height          =   285
         Left            =   -72720
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtrodf 
         Height          =   285
         Left            =   -72720
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.MonthView Calendar 
         Height          =   2370
         Left            =   5280
         TabIndex        =   5
         Top             =   1800
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         ShowToday       =   0   'False
         StartOfWeek     =   48955394
         CurrentDate     =   39264
      End
      Begin VB.TextBox lnprik 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "№ прод. приказа"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Размер шрифта для подрода войск"
         Height          =   615
         Left            =   -74760
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Размер шрифта для рода войск"
         Height          =   495
         Left            =   -74760
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "№ приказа по жетонам"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Menu mnuPaste 
      Caption         =   "Вставить"
      Visible         =   0   'False
      Begin VB.Menu muNum 
         Caption         =   "# - Порядковый &номер"
      End
      Begin VB.Menu mnuOblKom 
         Caption         =   "$ - Областная &команда"
      End
      Begin VB.Menu mnuDay 
         Caption         =   "# - &Число"
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "# - &Месяц"
      End
      Begin VB.Menu mnuYear 
         Caption         =   "# - &Год"
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar_dblclick()
Dim data As String
data = Calendar.Year & "-" & Format(Calendar.Month, "00") & "-" & Format(Calendar.Day, "00")
On Error GoTo a
Call mysql.query("Select ln, prod from dayprikaz where `data` = '" & data & "'")
prodprik.Text = DAT(2, 1)
lnprik.Text = DAT(1, 1)

GoTo b

a:
prodprik.Text = ""
lnprik.Text = ""

b:
End Sub

Private Sub Change_Click()
Dim Ln As String
Dim prod As String
Dim data As String
Dim VAldata As String

Ln = lnprik.Text
prod = prodprik.Text
data = Format(Calendar.VAlue, "YYYY-MM-DD")
Call mysql.query("Select count(*) from dayprikaz where `data` = '" & data & "'")
VAldata = DAT(1, 1)
If VAldata > 0 Then
Call mysql.query("UPDATE dayprikaz set `ln` = " & Ln & " ,`prod`= " & prod & " where `data` = '" & data & "'")
Else
Call mysql.query("INSERT into dayprikaz set `data` = '" & data & "', `ln` = " & Ln & " ,`prod`= " & prod & "")
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Call Change_Click
Unload Me
End Sub

Private Sub Form_Load()
Dim data As String

Calendar.VAlue = Format(Now, "DD.MM.YYYY")
On Error GoTo a
data = Format(Calendar.VAlue, "YYYY-MM-DD")
Call mysql.query("Select ln,prod from dayprikaz where `data` = '" & data & "'")
prodprik.Text = DAT(2, 1)
lnprik.Text = DAT(1, 1)
a:
End Sub
