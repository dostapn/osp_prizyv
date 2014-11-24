VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Дата"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7185
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdata.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Расширенный выбор"
      Height          =   3075
      Left            =   4860
      TabIndex        =   10
      Top             =   1125
      Width           =   2205
      Begin VB.TextBox DATE_txtCustom 
         Appearance      =   0  'Плоска
         BorderStyle     =   0  'Нет
         Enabled         =   0   'False
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   465
         TabIndex        =   3
         Top             =   1245
         Width           =   1575
      End
      Begin VB.CheckBox DATE_chCustom 
         Appearance      =   0  'Плоска
         Caption         =   "По шаблону:"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         TabIndex        =   2
         Top             =   885
         Width           =   1815
      End
      Begin VB.CheckBox DATE_chAll 
         Appearance      =   0  'Плоска
         Caption         =   "Все"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   0
         Top             =   330
         Width           =   1845
      End
      Begin VB.CheckBox DATE_chYear 
         Appearance      =   0  'Плоска
         Caption         =   "За текущий год"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblPrim 
         BackStyle       =   0  'Прозрачно
         Caption         =   $"frmdata.frx":08CA
         Height          =   1275
         Left            =   195
         TabIndex        =   11
         Top             =   1635
         Width           =   1815
      End
      Begin VB.Shape ShapeUpk 
         BackStyle       =   1  'Непрозрачно
         BorderColor     =   &H00FF8080&
         Height          =   300
         Left            =   375
         Top             =   1200
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   345
      Left            =   5955
      TabIndex        =   5
      Top             =   4635
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   60
      TabIndex        =   9
      Top             =   4455
      Width           =   7155
   End
   Begin VB.PictureBox picDate 
      Align           =   1  'Привязать вверх
      Appearance      =   0  'Плоска
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7155
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   7185
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Выберите дату комплектования комманд"
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
         Left            =   750
         TabIndex        =   8
         Top             =   270
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6240
         Picture         =   "frmdata.frx":0921
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSetDate 
      Caption         =   "Oк"
      Default         =   -1  'True
      Height          =   345
      Left            =   4830
      TabIndex        =   4
      Top             =   4635
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   120
      TabIndex        =   6
      Top             =   1125
      Width           =   4710
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Cal1_BeforeUpdate(Cancel As Integer)
On Error Resume Next
    DATE_chAll.Value = 0
    DATE_chCustom.Value = 0
    DATE_chYear.Value = 0
End Sub

Private Sub Cal1_Click()
    Caption = Cal1.Day & "." & MonthName(Cal1.Month, True) & "." & Cal1.year
End Sub

Private Sub Cal1_DblClick()
    cmdSetDate_Click
End Sub

Private Sub Cal1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 27 Then Unload Me
    End Sub

Private Sub cmdCancel_Click()
    Unload Me
    'frmTest.Show
End Sub

Private Sub cmdSetDate_Click()
On Error Resume Next
If DATE_chAll.Value = 1 Then
   strDATE = "%"
    frmListComm.d_DATE = "%"
    GoTo unl
End If

If DATE_chYear.Value = 1 Then
   strDATE = Format$(year(Date), , "00") & "-%"
    frmListComm.d_DATE = Format$(year(Date), , "00") & "-%"
    GoTo unl
End If

If DATE_chCustom.Value = 1 Then
   strDATE = DATE_txtCustom
    frmListComm.d_DATE = DATE_txtCustom
    GoTo unl
End If





    strDATE = Cal1.year & "-" & Format(Cal1.Month, "00") & "-" & Format(Cal1.Day, "00")
    frmListComm.d_DATE = Cal1.Day & "." & MonthName(Cal1.Month, True) & "." & Cal1.year
    Unload Me
'    If Not frmListComm.Visible Then frmListComm.Show vbModal, Me
unl:
   Unload Me
End Sub

Private Sub DATE_chAll_Click()
If DATE_chAll.Value = 1 Then
    DATE_chCustom.Value = 0
    DATE_chYear.Value = 0
End If
End Sub

Private Sub DATE_chCustom_Click()
 If DATE_chCustom.Value = 1 Then
    DATE_chAll.Value = 0
    DATE_chYear.Value = 0
End If

DATE_txtCustom.Enabled = DATE_chCustom.Value = 1
End Sub

Private Sub DATE_chYear_Click()
 If DATE_chYear.Value = 1 Then
    DATE_chCustom.Value = 0
    DATE_chAll.Value = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

End Sub

