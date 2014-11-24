VERSION 5.00
Begin VB.Form frmdelprnik 
   Caption         =   "Удаление призывника"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "frmdelprnik.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Удалить"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtprim 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   4815
   End
   Begin VB.ComboBox lstdel 
      BackColor       =   &H00C0FFC0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   0
      EndProperty
      Height          =   315
      ItemData        =   "frmdelprnik.frx":08CA
      Left            =   120
      List            =   "frmdelprnik.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmdelprnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim del As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOk_Click()
Dim comment As String
Dim in_com As String
Dim table_do As String
Dim Cont As Boolean
comment = lstdel.Text & " " & txtprim.Text
    Call Check_Sec
    If del Then
            If lstdel.Text = "Не прибыл" Then
                table_do = "delprnik"
                Cont = True
            End If
            If lstdel.Text = "Возврат" Then
                table_do = "vozvrat"
                    If Len(txtprim.Text) > "2" Then
                        Cont = True
                    Else
                        Exit Sub
                    End If
            End If
            If lstdel.Text = "Сбежал" Then
                table_do = "sochi"
                Cont = True
            End If
        End If
        
        If Cont = True Then
        Call mysql.query("SELECT * FROM prnik_" & nowBase & " WHERE idprnik='" & expupk & "'")
            For x = 1 To UBound(DAT())
                    If x = 25 Then DAT(25, 1) = DAT(25, 1) & comment
                    in_com = in_com & "'" & DAT(x, 1) & "'" & ","
            Next x
                in_com = in_com & "'" & lgn & "','" & CnvDataWinToSql(Date) & "'"
                Call mysql.query("insert into " & table_do & "_" & nowBase & " VAlues(" & in_com & ")")
                Call mysql.query("DELETE FROM prnik_" & nowBase & " WHERE `idprnik`='" & expupk & "'")
                Call log_sql("0", "3", expupk, lstdel.Text)
                MsgBox "Призывник был перенесен в базу '" & lstdel.Text & "' удачно!", vbInformation, "Удаление из базы."
        End If
        
        Call frmSearch.cmdSearch_Click
        Unload frmdelprnik
        Unload frmInfoPr
       
        

End Sub

Private Sub Form_Load()
On Error Resume Next
lstdel.AddItem ("Не прибыл")
lstdel.AddItem ("Возврат")
lstdel.AddItem ("Сбежал")
Call mysql.query("SELECT fam,name,otch,txtvk FROM prnik_" & nowBase & " WHERE idprnik='" & expupk & "'")
Label1.Caption = "Призывник " & DAT(1, 1) & " " & DAT(2, 1) & " " & DAT(3, 1) & ", " & DAT(4, 1) & " военкомат"
lstdel.ListIndex = 1
End Sub
Private Sub Check_Sec()
On Error Resume Next

Dim comm As String
Dim lock_pr As Long
Dim lprim As String

Call mysql.query("SELECT `lock`,lprim FROM prnik_" & nowBase & " WHERE idprnik='" & expupk & "'")
    lock_pr = DAT(1, 1)
    lprim = DAT(2, 1)
    'Нету блокировки: переносим призывника в таблицу удаленных
    If lock_rp = 0 Then
        If MsgBox("Вы уверены что хотите удалить этого призывника", vbOKCancel, "Удаление призывника") = vbOK Then
              del = True
              Exit Sub
              End If
    End If
    If lock_pr = 1 Then
              If MsgBox("Этот призывник заблокирован обычной блокировкой! Причина: " & lprim & ". Для удаления его нажмите <OK>.", vbOKCancel, "Удаление призывника") = vbOK Then
              del = True
              Exit Sub
              End If
    End If
    'Командная блокировка: разблокируем команду и потом токо удалять. Для
    'админов возможно просто удаить
    If lock_pr = 2 Or lock_pr = 4 Then
            If acl = "G" Then
                If MsgBox("Этот призывник находится в команде! Для удаления его нажмите <OK>.", vbOKCancel, "Удаление призывника") = vbOK Then
                    del = True
                    Exit Sub
                Else
                    MsgBox "Этот призывник находится в команде!Его удалить нельзя!!!", vbOK, "Удаление призывника"
                    Exit Sub
                End If
            End If
        
     End If
    'Блокировка администратора: снять может токо человек который входит в группу администраторов!!!
    
   If lock_pr = 3 And acl = "G" Then
        If MsgBox("Этот призывник заблокирован администратором! Причина: " & lprim & ". Для удаления его нажмите <OK>.", vbOKCancel, "Удаление призывника") = vbOK Then
        del = True
        Exit Sub
                Else
                    MsgBox "Этот призывник заблокирован администратором!Его удалить нельзя!!!", vbOK, "Удаление призывника"
                Exit Sub
        End If
    End If
End Sub
Private Sub DELETE_prnik()
        'Call del_prnik(comment)
        Call frmSearch.cmdSearch_Click
        Unload frmInfoPr
        Unload Me

End Sub
Private Sub del_prnik(Comments As String)
Dim in_com As String
Dim FIO As String
Call mysql.query("SELECT * FROM prnik_" & nowBase & " WHERE `idprnik`='" & expupk & "'")
FIO = DAT(3, 1) & " " & DAT(4, 1) & " " & DAT(5, 1)
For x = 1 To UBound(DAT())
    in_com = in_com & "'" & DAT(x, 1) & "'" & ","
Next x
in_com = in_com & "'" & Date & "','" & lgn & "','" & Comments & "'"
Call mysql.query("insert into delprnik_" & nowBase & " VAlues(" & in_com & ")")
Call mysql.query("DELETE FROM prnik_" & nowBase & " WHERE `idprnik`='" & expupk & "'")
in_com = vbNullString
Call mysql.query("SELECT * FROM prnik_" & nowBase & " WHERE `idprnik`='" & expupk & "'")
For x = 1 To UBound(DAT()) - 1
    in_com = in_com & "'" & DAT(x, 1) & "'" & ","
Next x
in_com = in_com & "'" & DAT(UBound(DAT()), 1) & "'"
Call mysql.query("insert into prnik_del_" & nowBase & " VAlues(" & in_com & ")")
Call mysql.query("DELETE FROM prnik_" & nowBase & " WHERE `idprnik`='" & expupk & "'")

MsgBox "Призывник " & FIO & " был удален из базы.", vbOKOnly, "Удаление призывника из базы."
End Sub
