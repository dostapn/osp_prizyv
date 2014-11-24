VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdir_dirs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Добавление директивы"
   ClientHeight    =   5745
   ClientLeft      =   3195
   ClientTop       =   2550
   ClientWidth     =   5790
   Icon            =   "frmdir_dirs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5790
   Begin MSComCtl2.DTPicker txtdirdate 
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24838145
      CurrentDate     =   39292
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "Добавить"
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtoblkom 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2500
      TabIndex        =   8
      Top             =   3480
      Width           =   3000
   End
   Begin VB.TextBox txtvch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2500
      TabIndex        =   6
      Top             =   2640
      Width           =   3000
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2500
      TabIndex        =   3
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label Label7 
      Caption         =   "Дата команды"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lddate 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Областная команда"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "Воинская часть"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "Дата"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2000
   End
   Begin VB.Label Label3 
      Caption         =   "Номер директивы"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2000
   End
   Begin VB.Label lID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2500
      TabIndex        =   1
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "ID директивы"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2000
   End
End
Attribute VB_Name = "frmdir_dirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idd As Long
Dim rodv As String
Public edit_pr As Boolean

Private Sub cmdAct_Click()
On Error Resume Next
If edit_pr = False Then
    If Len(txtDir) > 5 And Len(txtvch) > 3 And Len(txtoblkom) > 0 And Len(lddate) > 0 Then
        If mysql.query("insert into `directivi_" & nowBase & "` (`id`,`dir`,`data`,`vch`,`kom`,`kom_data`,`rodv`) values ('" & lID & "','" & txtDir & "','" & CnvDataWinToSql(txtdirdate) & "','" & txtvch & "','" & txtoblkom & "','" & CnvDataWinToSql(lddate) & "','" & rodv & "')") Then
            MsgBox "Директва удачно добавлена в базу", vbInformation, "Директивы"
            Unload Me
        Else
            MsgBox "Ошибка при добавлении директвы в базу", vbInformation, "Директивы"
        End If
    Else
        MsgBox "Ошибка при добавлении директвы в базу" & Chr(10) & "Вы ввели не все данные!!!", vbCritical, "Директивы"
    End If
Else
    If Len(txtDir) > 5 And Len(txtvch) > 3 And Len(txtoblkom) > 0 And Len(lddate) > 0 Then
        Call mysql.query("delete from `directivi_" & nowBase & "` where `id`='" & lID & "'")
        If mysql.query("insert into `directivi_" & nowBase & "` (`id`,`dir`,`data`,`vch`,`kom`,`kom_data`,`rodv`) values ('" & lID & "','" & txtDir & "','" & CnvDataWinToSql(txtdirdate) & "','" & txtvch & "','" & txtoblkom & "','" & CnvDataWinToSql(lddate) & "','" & rodv & "')") Then
            MsgBox "Директва удачно обновлена", vbInformation, "Директивы"
            Unload Me
        Else
            MsgBox "Ошибка при обновлении директвы", vbInformation, "Директивы"
        End If
    Else
        MsgBox "Ошибка при добавлении директвы в базу" & Chr(10) & "Вы ввели не все данные!!!", vbCritical, "Директивы"
    End If

End If
Call frmdir.load_dirs
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()

If edit_pr = True Then Call edit_mode
Call mysql.query("select max(id) from directivi_" & nowBase)
idd = val(DAT(1, 1)) + 1
lID.Caption = idd
If edit_pr = True Then edit_mode: Exit Sub
End Sub
Private Sub edit_mode()
On Error Resume Next
Call mysql.query("select `id`,`dir`,`data`,`vch`,`kom` from directivi_" & nowBase & " where `id`='" & frmdir.lstdirs.SelectedItem.Text & "'")
lID.Caption = DAT(1, 1)
txtDir = DAT(2, 1)
txtdirdate = DAT(3, 1)
txtvch = DAT(4, 1)
txtoblkom = DAT(5, 1)
cmdAct.Caption = "Обновить"
Me.Caption = "Обновление директивы"
Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call frmdir.sum_dir
End Sub

Private Sub txtoblkom_Change()
On Error Resume Next
Label2.Caption = vbNullString
Call mysql.query("SELECT `datanar`,`rodv` from naryad_" & nowBase & " where `oblkom`='" & txtoblkom & "'")
If st > 0 Then
    lddate.Caption = CnvDataSqLToWin(DAT(1, 1)): rodv = DAT(2, 1)
Else
    lddate.Caption = vbNullString
End If
End Sub
