VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFullReport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Отчет"
   ClientHeight    =   8145
   ClientLeft      =   -555
   ClientTop       =   15
   ClientWidth     =   6240
   HasDC           =   0   'False
   Icon            =   "frmFullReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker txtDatePo 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49217537
      CurrentDate     =   39248
   End
   Begin MSComCtl2.DTPicker txtDateS 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648384
      Format          =   49217539
      CurrentDate     =   39248
   End
   Begin VB.CommandButton cmdvod 
      Caption         =   "Водители"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Обновить"
      Height          =   480
      Left            =   3720
      Picture         =   "frmFullReport.frx":08CA
      TabIndex        =   2
      Top             =   120
      Width           =   1140
   End
   Begin MSComctlLib.ListView listfiles 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   12091
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Военкомат"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "К-во УПК"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2275
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22.11.2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:51"
         EndProperty
      EndProperty
   End
   Begin VB.Menu to_excel 
      Caption         =   "В Excel"
   End
End
Attribute VB_Name = "frmFullReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdRefresh_Click()
    Call RefList
End Sub
Private Sub cmdvod_Click()
Dim nC As Long
Dim x As Long
Dim st As Long
Dim tVK As String
Dim vv As Long
listfiles.ListItems.Clear
Call Reg_VK_List
For vv = 0 To UBound(nVK())
Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `dataosp` >= '" & CnvDataWinToSql(txtDateS) & "' AND `dataosp` <= '" & CnvDataWinToSql(txtDatePo) & "' AND `txtvk` like '%" & nVK(vv) & "%' AND vod='1' AND otprvid>'0'")
    Set LF = listfiles.ListItems.add(, , nVK(vv))
    LF.SubItems(1) = DAT(1, 1)
    If DAT(1, 1) = "0" Then LF.ListSubItems.Item(1).ForeColor = &HFF&
    st = DAT(1, 1) + st
Next vv
listfiles.Refresh
SB1.Panels(1).Text = "Всего: " & st
End Sub
Private Sub Form_Load()
    txtDateS = Date
    txtDatePo = Date
    Call RefList
End Sub
Sub RefList()
Dim nC As Long
Dim x As Long
Dim st As Long
Dim tVK As String
Dim vv As Long
listfiles.ListItems.Clear
Call Reg_VK_List
For vv = 0 To UBound(nVK())
Call mysql.query("SELECT Count(idprnik) FROM prnik_" & nowBase & " WHERE `dataosp` >= '" & CnvDataWinToSql(txtDateS) & "' AND `dataosp` <= '" & CnvDataWinToSql(txtDatePo) & "' AND `txtvk` like '%" & nVK(vv) & "%'")
    Set LF = listfiles.ListItems.add(, , nVK(vv))
    LF.SubItems(1) = DAT(1, 1)
    If DAT(1, 1) = "0" Then LF.ListSubItems.Item(1).ForeColor = &HFF&
    st = DAT(1, 1) + st
Next vv
listfiles.Refresh
SB1.Panels(1).Text = "Всего: " & st
End Sub
Private Sub to_excel_Click()
Call Cnv.ResoultSearch(listfiles, True, "ListCommAND", Caption)
End Sub
