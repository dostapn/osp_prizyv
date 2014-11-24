VERSION 5.00
Begin VB.Form frmdelinfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Востановление призывника"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   FillColor       =   &H0000FFFF&
   FillStyle       =   4  'Upward Diagonal
   Icon            =   "frmdelinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdprint 
      Caption         =   "Печать"
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Сохранить"
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   6240
      Width           =   1245
   End
   Begin VB.TextBox datedel 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox lprim 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Восстановить"
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   6240
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ОК"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1125
   End
   Begin VB.Label whodel 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   13
      Top             =   3840
      Width           =   3000
   End
   Begin VB.Label vk 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   3000
   End
   Begin VB.Label datar 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   11
      Top             =   2040
      Width           =   3000
   End
   Begin VB.Label Otch 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   3000
   End
   Begin VB.Label name_b 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   3000
   End
   Begin VB.Label fam 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Причина"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Кем удалено"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Дата удаления из базы"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Военкомат"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Дата рождения"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Отчество"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Имя"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Фамилия"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1900
   End
End
Attribute VB_Name = "frmdelinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
On Error Resume Next
''Screen.MousePointer = vbHourglass
    Set oExcelApp = CreateObject("EXCEL.APPLICATION")
    Dim sFile As String
    sFile = sCNV_txtDirShabl & "info_d.xls"
    If iCNV_chShowObj = 1 Then oExcelApp.Visible = True
    oExcelApp.WORkbooks.Open FileName:=sFile, ReadOnly:=True, ignOReReadOnlyRecommended:=True
    
    Set oWb = oExcelApp.ActiveWORkbook
    Set oWs = oExcelApp.Sheets(1)
Call mysql.query("SELECT idprnik,fam,name,otch,datar,txtobraz,txtspec,txtvk,dataosp,servb,nomvb,txtsem,txtsud,vod,lprim,datedel FROM " & table & "_" & nowBase & " WHERE idprnik = '" & expupk & "'")
With oWs
          
                       
            .cells(1, 2) = DAT(1, 1)
            .cells(3, 2) = DAT(2, 1)
            .cells(4, 2) = DAT(3, 1)
            .cells(5, 2) = DAT(4, 1)
            .cells(6, 2) = DAT(5, 1)
            .cells(8, 2) = DAT(6, 1)
            .cells(9, 2) = DAT(7, 1)
            .cells(11, 2) = DAT(8, 1)
            .cells(12, 2) = DAT(9, 1)
            .cells(13, 2) = DAT(10, 1) & " " & DAT(11, 1)
            .cells(15, 2) = DAT(12, 1)
            .cells(16, 2) = DAT(13, 1)
            
            If DAT(14, 1) = 1 Then
                .cells(17, 2) = "ДА"
            Else
                .cells(17, 2) = "НЕТ"
            End If
            
            .cells(19, 2) = DAT(15, 1)
            .cells(20, 2) = DAT(16, 1)
        
            
End With
End Sub

Private Sub cmdsave_Click()
If Len(CnvDataWinToSql(datedel.Text)) > 0 And Len(lprim.Text) > 0 Then
    Call mysql.query("UPDATE `" & table & "_" & nowBase & "` set datedel='" & CnvDataWinToSql(datedel.Text) & "' WHERE idprnik='" & expupk & "'")
    Call mysql.query("UPDATE `" & table & "_" & nowBase & "` set lprim='" & lprim.Text & "' WHERE idprnik='" & expupk & "'")
End If
End Sub

Private Sub CommAND1_Click()
Unload Me
End Sub

Private Sub CommAND2_Click()
Dim newDate As String
Dim newUpk As Integer
Dim inf As String
Dim comment As String
If MsgBox("Восстановить призывника?", vbYesNo + vbQuestion, strMAIN_TITLE) = vbYes Then
    
    
          If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
        nUpk = Int(frmSearch.listres.ListItems(frmSearch.listres.SelectedItem.Index).Text)
        frmSearch.listres.ListItems.Remove (frmSearch.listres.SelectedItem.Index)
      
    End If
    
        newDate = CnvDataWinToSql(Date)
        Call mysql.query("SELECT * FROM prnik_" & nowBase & " WHERE idprnik='" & nUpk & "'")
             If st > 0 Then
                Call mysql.query("SELECT max(idprnik) FROM prnik_" & nowBase & "")
                newUpk = DAT(1, 1) + 1
        Else
            newUpk = nUpk
        End If
        
        Call mysql.query("SELECT * FROM `" & table & "_" & nowBase & "` WHERE idprnik=" & nUpk)
            in_com = "'" & (newUpk) & "'" & ","
            For x = 2 To (UBound(DAT()) - 3)
            comment = "Восстановлен. (" & DAT(25, 1) & ")"
                    If x = 25 Then DAT(25, 1) = comment
                    in_com = in_com & "'" & DAT(x, 1) & "'" & ","
            Next x
                in_com = in_com & "'" & DAT((UBound(DAT()) - 2), 1) & "'"
         Call mysql.query("INSERT INTO `prnik_" & nowBase & "` VALUES (" & in_com & ") ")
        
        inf = DAT(3, 1) & " " & DAT(4, 1) & " " & DAT(5, 1) & " " & DAT(2, 1)
        Call mysql.query("DELETE FROM `" & table & "_" & nowBase & "` WHERE idprnik=" & nUpk)
        Call log_sql("0", "4", expupk, "Востановлен")
                  MsgBox "Призывник удачно востановлен!!!", vbInformation, "Востановление"
Unload Me
Call frmSearch.cmdSearch_Click
End Sub
Private Sub Form_Load()
On Error Resume Next
Call mysql.query("SELECT fam,name,otch,datar,txtvk,datedel,who,lprim FROM " & table & "_" & nowBase & " WHERE idprnik = '" & expupk & "'")
fam.Caption = DAT(1, 1)
name_b.Caption = DAT(2, 1)
Otch.Caption = DAT(3, 1)
datar.Caption = CnvDataSqLToWin(DAT(4, 1))
vk.Caption = DAT(5, 1)
datedel.Text = CnvDataSqLToWin(DAT(6, 1))
lprim.Text = DAT(8, 1)
Call mysql.query("SELECT fio FROM users WHERE name='" & DAT(7, 1) & "'")
whodel.Caption = DAT(1, 1)
End Sub
