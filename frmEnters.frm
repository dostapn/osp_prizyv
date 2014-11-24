VERSION 5.00
Begin VB.Form frmEnter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ОСП Призыв 2.0 Воловин Edition"
   ClientHeight    =   2655
   ClientLeft      =   7185
   ClientTop       =   6675
   ClientWidth     =   3960
   ClipControls    =   0   'False
   Icon            =   "frmEnters.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton enter 
      BackColor       =   &H80000016&
      Caption         =   "Вход"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox passwd 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox login 
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
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Пароль"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Логин"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub asxBars1_Click()

End Sub
Private Sub enter_Click()
On Error Resume Next
Dim ddate As Date


lgn = login.Text
If Len(lgn) > 0 And Len(passwd.Text) > 0 Then
    Call cmdRCon_Click
    Call mysql.query("SELECT pass FROM users WHERE name='" & lgn & "'")
        If st > 0 Then
            If passwd = DAT(1, 1) Then bd_access = True
                If bd_access Then
                    Unload Me
                        frmMain.Show
              
                Else
                    MsgBox "Не правильно введен пароль!!!", vbCritical, "ОШИБКА"
                    login.Text = vbNullString
                    passwd.Text = vbNullString
                    login.SetFocus
                End If
            Else
                    MsgBox "Не правильно введен логин/пароль!!!", vbCritical, "ОШИБКА"
                    login.Text = vbNullString
                    passwd.Text = vbNullString
                    login.SetFocus
            End If
        
        End If
        
End Sub

Private Sub Form_Load()

Call INIT_VRP_LIST
Call READ_CONFIG
        
        txtHost = H_NAME
        txtUser = U_NAME
        txtpass = C_PASSWORD
        txtDB = D_BASE
        txtPORt = C_PORT

  End Sub

Private Sub cmdRCon_Click()
On Error GoTo errhANDler
Dim choose As Boolean
connect:

        DoEvents
        Set mysql = New cMysql
        Call mysql.real_connect(txtHost, txtUser, txtpass, txtDB)
        Exit Sub
GoTo connect
errhANDler:
        Screen.MousePointer = vbDefault
         Caption = Err.Description
End Sub

