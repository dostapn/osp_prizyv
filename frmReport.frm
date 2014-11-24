VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Отчет за день"
   ClientHeight    =   5250
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ок"
      Default         =   -1  'True
      Height          =   330
      Left            =   4080
      TabIndex        =   18
      Top             =   4845
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   330
      Left            =   5130
      TabIndex        =   17
      Top             =   4845
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Количество"
      Height          =   3030
      Left            =   315
      TabIndex        =   2
      Top             =   1575
      Width           =   5625
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Обновить"
         Height          =   360
         Left            =   1680
         TabIndex        =   21
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CommandButton cmdRay 
         Caption         =   "Районка..."
         Height          =   360
         Left            =   2940
         TabIndex        =   19
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CommandButton cmdAd 
         Caption         =   "Подробно..."
         Height          =   360
         Left            =   4200
         TabIndex        =   15
         Top             =   2460
         Width           =   1215
      End
      Begin VB.TextBox txtReg 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   2730
      End
      Begin VB.TextBox txtDel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   705
         Width           =   2730
      End
      Begin VB.TextBox txtSend 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   1050
         Width           =   2730
      End
      Begin VB.TextBox txtOst 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1395
         Width           =   2730
      End
      Begin VB.TextBox txtVod 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1740
         Width           =   2730
      End
      Begin VB.TextBox txtDir 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   2085
         Width           =   2730
      End
      Begin VB.Shape ShapeUpk 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   315
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   660
         Width           =   2865
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   1005
         Width           =   2865
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   1350
         Width           =   2865
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Зарегистрированных:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   390
         Width           =   1680
      End
      Begin VB.Label lblDir 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Директивщиков:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1125
         TabIndex        =   13
         Top             =   2085
         Width           =   1305
      End
      Begin VB.Label lblDel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Удаленных:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1500
         TabIndex        =   12
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lblVod 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Водителей:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1545
         TabIndex        =   11
         Top             =   1755
         Width           =   885
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   1695
         Width           =   2865
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   2550
         Top             =   2040
         Width           =   2865
      End
      Begin VB.Label lblSend 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Отправленных:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1230
         TabIndex        =   10
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label lblOst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Осталось на Сборном пункте:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1410
         Width           =   2280
      End
   End
   Begin VB.PictureBox picReport 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   6210
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1500
         TabIndex        =   20
         Text            =   "09.11.2004"
         Top             =   300
         Width           =   2625
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   1410
         Top             =   255
         Width           =   2760
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5430
         Picture         =   "frmReport.frx":08CA
         Top             =   150
         Width           =   480
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Сведения на "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   315
         Width           =   1155
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3630
      Left            =   150
      TabIndex        =   16
      Top             =   1140
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   6403
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Сведения"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAd_Click()
    frmFullReport.Show vbModal, frmMain
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRay_Click()
    Call Cnv.To_Ray
End Sub

Private Sub cmdRefresh_Click()
    Call RefList
End Sub

Private Sub Form_Load()
txtDate = Date

Call RefList

End Sub

Sub RefList()
On Error GoTo ErrH
    Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `dataosp` = '" & CnvDataWinToSql(txtDate) & "'")
    txtReg = DAT(1, 1)
    
    

    Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `otprvid` <> 0 AND prnik_" & nowBase & ".dataosp = '" & CnvDataWinToSql(txtDate) & "'")
    txtSend = DAT(1, 1)
    
    
    Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `otprvid` = 0 AND prnik_" & nowBase & ".dataosp = '" & CnvDataWinToSql(txtDate) & "'")
    txtOst = DAT(1, 1)
    
    
        Call mysql.query("SELECT Count(vod) FROM prnik_" & nowBase & " WHERE `vod` = 1 AND `dataosp` like '" & CnvDataWinToSql(txtDate) & "%'")
    txtVod = DAT(1, 1)
    
        Call mysql.query("SELECT Count(dir) FROM prnik_" & nowBase & " WHERE `dir` <> ''  AND `dataosp` like '" & CnvDataWinToSql(txtDate) & "%'")
    txtDir = DAT(1, 1)
Exit Sub
ErrH:
MsgBox "Непредвиденная ошибка" & NL2 & Err.Description, vbCritical, strMAIN_TITLE

End Sub


