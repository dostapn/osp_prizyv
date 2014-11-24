VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnaryad_com_add 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Добавление команды"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt3 
      Caption         =   "ЗАТО"
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   7320
      Width           =   975
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Подшефные части"
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   7320
      Width           =   1815
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Обычная"
      Height          =   375
      Left            =   200
      TabIndex        =   19
      Top             =   7320
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "Добавить"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker txtdatanar 
      Height          =   300
      Left            =   2685
      TabIndex        =   7
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      _Version        =   393216
      CalendarBackColor=   8454016
      Format          =   49086465
      CurrentDate     =   39267
   End
   Begin VB.TextBox txtkolvo 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      TabIndex        =   6
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtpred 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      TabIndex        =   5
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtvch 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      TabIndex        =   4
      Top             =   5040
      Width           =   2295
   End
   Begin VB.ComboBox txtdoroga 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2685
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtpunkt 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.ComboBox txtvokrug 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2685
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtokr 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2685
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "Подрод войск"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Дата наряда"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   17
      Top             =   6720
      Width           =   1995
   End
   Begin VB.Label Label10 
      Caption         =   "Колитчество"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   16
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label Label9 
      Caption         =   "Предназначение"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   15
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Label Label8 
      Caption         =   "Воинскач часть"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   14
      Top             =   4920
      Width           =   1995
   End
   Begin VB.Label Label7 
      Caption         =   "Дорога"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   13
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "Пункт"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   12
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "Военный округ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   11
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Окружная команда"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   10
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   200
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmnaryad_com_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public major, minor As Integer
Dim okr() As String
Dim okr_e As Integer

Private Sub cmd_add_Click()
'predpodgotovka
On Error Resume Next
Dim vch_type, obl As Integer
Dim lastk, lastk2, rodv, oblkom As String
'наличие всех данных

If Not IsNumeric(txtkolvo.Text) Then
    MsgBox "Вы указали неправильное количество призывников!", vbCritical, "Ошибка"
    Exit Sub
End If


If Len(txtokr.Text) Or Len(txtpunkt.Text) Or Len(txtpunkt.Text) Or Len(txtvch.Text) Or Len(txtpred.Text) Then
    Else
    MsgBox "Вы заполнили не все поля!", vbCritical, "Ошибка"
    Exit Sub
End If

'

If opt1 = True Then vch_type = "0"
If opt2 = True Then vch_type = "1"
If opt3 = True Then vch_type = "2"

Call mysql.query("SELECT name FROM naryad_rodv_" & nowBase & " WHERE minor='1' AND majOR='" & major & "'")
rodv = DAT(1, 1)
'поиск команды с таким же номером(окружным)
Call komend_in
okr = Split(txtokr, "/")
okr_e = okr(UBound(okr()))
Dim okr_db As String


If UBound(okr) - 1 > 0 Then
    okr_db = okr(0) & "/" & okr(1) & "/"
Else
    okr_db = okr(0) & "/"
End If



Call mysql.query("SELECT okrkom,oblkom,narid,okrkom_e FROM naryad_" & nowBase & " WHERE okrkom='" & okr_db & "' AND okrkom_e='" & okr_e & "'")
    
    If st > 0 Then
        lastk = DAT(2, st)
        If st = 1 Then
            If IsNumeric(lastk) Then
                lastk = lastk & komend(0)
                Call mysql.query("UPDATE naryad_" & nowBase & " set oblkom='" & lastk & "' WHERE narid='" & DAT(3, 1) & "'")
              End If
        End If

                Dim tmpp, tmpp2 As String
                Dim x, lastX As Integer
                
                    tmpp = Left$(lastk, Len(lastk) - 1) 'обрезаем последний символ
                    
                    If IsNumeric(tmpp) Then ' смотрим если еще буквы в названии команды
                        
                            ' смотрим не последняя ли буква
                            If Right(lastk, 1) = komend(UBound(komend())) Then
                                oblkom = tmpp & komend(0) & komend(0)
                            Else
                                For x = 0 To UBound(komend())  'запускаем поиск индекса с  нашей буквой
                                    If Right(lastk, 1) = komend(x) Then lastX = x: Exit For
                                Next x
                        oblkom = tmpp & komend(lastX + 1)
                            End If
                    Else 'если две буквы(надеюсь больше не понадобится!!!)
                           ' смотрим не последняя ли буква
                           If Right(lastk, 1) = komend(UBound(komend())) Then
                                
                                    For x = 0 To UBound(komend())  'запускаем поиск индекса с  нашей буквой
                                        If Right(tmpp, 1) = komend(x) Then lastX = x
                                    Next x
                                    
                                oblkom = Left(tmpp, Len(tmpp) - 1) & komend(lastX + 1) & komend(0)
                            Else
                            
                                 For x = 0 To UBound(komend())  'запускаем поиск индекса с  нашей буквой
                                        If Right(lastk, 1) = komend(x) Then lastX = x
                                Next x
                                
                                oblkom = tmpp & komend(lastX + 1)
                            End If
                            
                    End If
    Else
        Call mysql.query("SELECT max(oblkom) FROM naryad_" & nowBase)
        lastk = DAT(1, 1)
            If lastk = "" Then
                oblkom = "1"
            Else
            'если просто цифра
                    If IsNumeric(lastk) Then
                        lastk = lastk
                    Else
                        If IsNumeric(Left(lastk, Len(lastk) - 1)) Then
                            lastk = Left(lastk, Len(lastk) - 1)
                        Else
                            lastk = Left(lastk, Len(lastk) - 2)
                        End If
                    End If
                   
                    If Right(lastk, 1) = "0" Then
                         oblkom = lastk + 2
                     Else
                         oblkom = lastk + 1
                    End If
                
            End If
    End If

Call mysql.query("insert into naryad_" & nowBase & " ( `narid` , `okrkom` , `okrkom_e`,`oblkom` , `rodv` , `okr` , `punkt` , `doroga` , `vch` , `vrp` , `kolvo` , `datanar` , `major` , `minor` , `type` ) values ('" & lid.Caption & "','" & okr_db & "','" & okr_e & "','" & oblkom & "','" & rodv & "','" & txtvokrug.Text & "','" & txtpunkt.Text & "','" & txtdoroga.Text & "','" & txtvch.Text & "','" & txtpred.Text & "','" & txtkolvo.Text & "','" & CnvDataWinToSql(txtdatanar) & "','" & major & "','" & minor & "','" & vch_type & "')")

Call frmNaryad.rodvtr_Click
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim ID, x As Long
  'new id
    Call mysql.query("SELECT max(narid) FROM naryad_" & nowBase)
    If DAT(1, 1) = "" Then
        ID = "1"
    Else
        ID = DAT(1, 1) + 1
    End If
    
    lid = ID
  'end new id
'input
If minor = "1" Then
    txtdop.Caption = "Нет"
    Call mysql.query("SELECT `name` FROM naryad_rodv_" & nowBase & " WHERE majOR='" & major & "' AND minor='1'")
    txtosn.Caption = DAT(1, 1)
Else
    Call mysql.query("SELECT `name` FROM naryad_rodv_" & nowBase & " WHERE majOR='" & major & "' AND minor='1'")
    txtosn.Caption = DAT(1, 1)
    Call mysql.query("SELECT `name` FROM naryad_rodv_" & nowBase & " WHERE majOR='" & major & "' AND minor='" & minor & "'")
    txtdop.Caption = DAT(1, 1)
End If



txtvokrug.AddItem ("БФ")
txtvokrug.AddItem ("ДВО")
txtvokrug.AddItem ("КВ")
txtvokrug.AddItem ("ЛенВО")
txtvokrug.AddItem ("МВО")
txtvokrug.AddItem ("ПУрВО")
txtvokrug.AddItem ("СибВО")
txtvokrug.AddItem ("СФ")
txtvokrug.AddItem ("ТОФ")
txtvokrug.AddItem ("ЧФ")

txtdoroga.AddItem ("Самолет")
txtdoroga.AddItem ("Октябрьская")
txtdoroga.AddItem ("Калиниградская")
txtdoroga.AddItem ("Московская")
txtdoroga.AddItem ("Северная")
txtdoroga.AddItem ("Горьковская")
txtdoroga.AddItem ("Куйбышевская")
txtdoroga.AddItem ("Юго-Восточная")
txtdoroga.AddItem ("Приволжская")
txtdoroga.AddItem ("Северо-Кавказская")
txtdoroga.AddItem ("Свердловская")
txtdoroga.AddItem ("Южно-Уральская")
txtdoroga.AddItem ("Заподно-Сибирская")
txtdoroga.AddItem ("Красноярская")
txtdoroga.AddItem ("Восточно-Сибирская")
txtdoroga.AddItem ("Забайкальская")
txtdoroga.AddItem ("Дальневосточная")
txtdoroga.AddItem ("Сахалинская")

txtdatanar = Date
End Sub

Private Sub opt3_Click()
txtvch.Text = txtvch.Text & "(ЗАТО)"
End Sub

Private Sub txtokr_LostFocus()
On Error Resume Next
Dim okr() As String
okr = Split(txtokr, "/")
okr_e = okr(UBound(okr()))
Dim okr_db As String
If UBound(okr) - 1 > 0 Then
    okr_db = okr(0) & "/" & okr(1) & "/"
Else
    okr_db = okr(0) & "/"
End If
Call mysql.query("SELECT `okr`,`punkt`,`doroga`,`vch` FROM naryad_" & nowBase & " WHERE okrkom='" & okr_db & "' AND okrkom_e='" & okr_e & "'")
If st > 0 Then
    txtvokrug.Text = DAT(1, 1)
    txtpunkt.Text = DAT(2, 1)
    txtdoroga.Text = DAT(3, 1)
    txtvch.Text = DAT(4, 1)
End If
End Sub
